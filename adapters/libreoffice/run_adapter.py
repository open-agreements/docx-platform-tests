#!/usr/bin/env python3
"""docx-platform-tests adapter protocol v1 entrypoint for LibreOffice.

Capability map (see README.md):
- acceptAllTrackedChanges / rejectAllTrackedChanges: LibreOffice UNO dispatches.
- replaceFirstTextOccurrence: LibreOffice UNO search descriptor + text range
  replacement, using case-sensitive non-regex document-order search.
"""
import argparse
import json
import os
import signal
import shutil
import socket
import subprocess
import sys
import tempfile
import time
from pathlib import Path

PROTOCOL_VERSION = "1"
SUPPORTED_OPERATIONS = {
    "acceptAllTrackedChanges",
    "rejectAllTrackedChanges",
    "replaceFirstTextOccurrence",
}
DOCX_FILTER_NAME = "MS Word 2007 XML"
CONNECT_TIMEOUT_SECONDS = 20.0
OPERATION_TIMEOUT_SECONDS = 60
SHUTDOWN_TIMEOUT_SECONDS = 5.0


def find_free_port() -> int:
    with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as sock:
        sock.bind(("127.0.0.1", 0))
        return int(sock.getsockname()[1])


def uno_file_url(path: str) -> str:
    return Path(path).resolve().as_uri()


def property_value(uno_module, name: str, value):
    prop = uno_module.createUnoStruct("com.sun.star.beans.PropertyValue")
    prop.Name = name
    prop.Value = value
    return prop


def connect_to_soffice(uno_module, port: int, process: subprocess.Popen):
    local_context = uno_module.getComponentContext()
    resolver = local_context.ServiceManager.createInstanceWithContext(
        "com.sun.star.bridge.UnoUrlResolver", local_context
    )
    url = f"uno:socket,host=localhost,port={port};urp;StarOffice.ComponentContext"
    deadline = time.monotonic() + CONNECT_TIMEOUT_SECONDS
    last_error = None

    while time.monotonic() < deadline:
        if process.poll() is not None:
            raise RuntimeError(f"soffice exited before accepting UNO connections: {process.returncode}")
        try:
            return resolver.resolve(url)
        except Exception as exc:  # UNO raises implementation-specific exceptions here.
            last_error = exc
            time.sleep(0.2)

    raise RuntimeError(f"timed out connecting to soffice UNO listener: {last_error}")


def store_docx(uno_module, document, output_path: str) -> None:
    document.storeToURL(
        uno_file_url(output_path),
        (property_value(uno_module, "FilterName", DOCX_FILTER_NAME),),
    )


def apply_operation(service_manager, remote_context, document, operation: dict) -> None:
    name = operation.get("operationName")

    if name in ("acceptAllTrackedChanges", "rejectAllTrackedChanges"):
        frame = document.getCurrentController().getFrame()
        dispatcher = service_manager.createInstanceWithContext(
            "com.sun.star.frame.DispatchHelper", remote_context
        )
        command = (
            ".uno:AcceptAllTrackedChanges"
            if name == "acceptAllTrackedChanges"
            else ".uno:RejectAllTrackedChanges"
        )
        dispatcher.executeDispatch(frame, command, "", 0, ())
        return

    if name == "replaceFirstTextOccurrence":
        descriptor = document.createSearchDescriptor()
        descriptor.SearchString = operation["findText"]
        descriptor.SearchCaseSensitive = True
        descriptor.SearchRegularExpression = False
        descriptor.SearchWords = False
        found = document.findFirst(descriptor)
        if found is None:
            raise RuntimeError(f"findText not present in any paragraph: {operation['findText']!r}")
        found.setString(operation["replaceText"])
        return

    raise AssertionError(f"unsupported operation reached apply_operation: {name!r}")


def run_supported_operation(args, operation: dict) -> int:
    import uno  # imported late so unsupported/mismatch paths need no pyuno

    soffice = shutil.which("soffice")
    if soffice is None:
        print("soffice not found on PATH", file=sys.stderr)
        return 1

    work_dir = tempfile.mkdtemp(prefix="docx-platform-lo-")
    profile_dir = os.path.join(work_dir, "profile")
    os.makedirs(profile_dir, exist_ok=True)
    port = find_free_port()
    process = None
    document = None

    try:
        process = subprocess.Popen(
            [
                soffice,
                "--headless",
                "--invisible",
                "--norestore",
                "--nologo",
                f"-env:UserInstallation={Path(profile_dir).resolve().as_uri()}",
                f"--accept=socket,host=localhost,port={port};urp;StarOffice.ComponentContext",
            ],
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
            text=True,
        )

        remote_context = connect_to_soffice(uno, port, process)
        service_manager = remote_context.ServiceManager
        desktop = service_manager.createInstanceWithContext(
            "com.sun.star.frame.Desktop", remote_context
        )
        load_props = (property_value(uno, "Hidden", True),)
        document = desktop.loadComponentFromURL(
            uno_file_url(args.input), "_blank", 0, load_props
        )
        if document is None:
            raise RuntimeError(f"LibreOffice failed to load input document: {args.input}")

        apply_operation(service_manager, remote_context, document, operation)
        store_docx(uno, document, args.output)
        return 0
    except Exception as exc:
        print(f"libreoffice adapter error: {exc}", file=sys.stderr)
        return 1
    finally:
        if document is not None:
            try:
                document.close(False)
            except Exception:
                try:
                    document.dispose()
                except Exception:
                    pass
        if process is not None:
            if process.poll() is None:
                process.terminate()
                try:
                    process.communicate(timeout=SHUTDOWN_TIMEOUT_SECONDS)
                except subprocess.TimeoutExpired:
                    process.kill()
                    process.communicate(timeout=SHUTDOWN_TIMEOUT_SECONDS)
            else:
                process.communicate(timeout=SHUTDOWN_TIMEOUT_SECONDS)
        shutil.rmtree(work_dir, ignore_errors=True)


def operation_timeout(_signum, _frame):
    raise TimeoutError(
        f"LibreOffice operation exceeded {OPERATION_TIMEOUT_SECONDS} seconds"
    )


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--protocol-version", required=True)
    parser.add_argument("--operation", required=True)
    parser.add_argument("--input", required=True)
    parser.add_argument("--output", required=True)
    args = parser.parse_args()

    if args.protocol_version != PROTOCOL_VERSION:
        print(
            f"libreoffice adapter speaks protocol v{PROTOCOL_VERSION}, "
            f"got {args.protocol_version}"
        )
        return 3

    with open(args.operation, encoding="utf-8") as f:
        operation = json.load(f)
    name = operation.get("operationName")

    if name not in SUPPORTED_OPERATIONS:
        print(f"libreoffice adapter does not implement operation '{name}'")
        return 2

    previous_handler = signal.signal(signal.SIGALRM, operation_timeout)
    signal.alarm(OPERATION_TIMEOUT_SECONDS)
    try:
        return run_supported_operation(args, operation)
    finally:
        signal.alarm(0)
        signal.signal(signal.SIGALRM, previous_handler)


if __name__ == "__main__":
    sys.exit(main())
