"""Move a date range of Microsoft 365 messages to another folder."""

import os
import sys
import time
from datetime import date, timedelta
from typing import Any
from urllib.parse import quote, urlencode

import msal
import requests


# ===== Settings to change ====================================================
# Create an Entra app registration, enable public-client flows, and grant its
# delegated Microsoft Graph permission: Mail.ReadWrite.Shared
TENANT_ID = os.environ["MS_TENANT_ID"]
CLIENT_ID = os.environ["MS_CLIENT_ID"]

# Sign in with an account delegated Mail.ReadWrite.Shared access to this mailbox.
MAILBOX = "buero@kkht.de"
# Folder paths are slash-separated names relative to the mailbox root. The
# first item may be a Graph well-known name, avoiding localization differences.
# Each pair is (source folder, Graph date field). Use "sentDateTime" for
# Sent Items and "receivedDateTime" for Inbox and most received-mail folders.
#SOURCES = [("sentitems", "sentDateTime")]
SOURCES = [("inbox", "receivedDateTime"), ("sentitems", "sentDateTime")]
DESTINATION_FOLDER = "archive/2025"

# Both dates are inclusive UTC calendar dates.
START_DATE = "2025-01-01"
END_DATE = "2025-01-30"

# Review the candidates first. Set False only when ready to perform moves.
DRY_RUN = True
# =============================================================================

GRAPH_ROOT = "https://graph.microsoft.com/v1.0"
SCOPES = ["https://graph.microsoft.com/Mail.ReadWrite.Shared"]
WELL_KNOWN_FOLDERS = frozenset(
    {"archive", "deleteditems", "drafts", "inbox", "junkemail", "outbox", "sentitems"}
)


def graph_request(
    url: str, token: str, method: str = "GET", body: dict[str, Any] | None = None
) -> dict[str, Any]:
    """Call Graph, retrying transient throttling and server failures."""
    headers = {"Authorization": f"Bearer {token}", "Accept": "application/json"}

    for attempt in range(5):
        try:
            # URLs are Graph endpoints or Graph-provided @odata.nextLink values.
            response = requests.request(method, url, headers=headers, json=body, timeout=60)
        except requests.RequestException as error:
            if attempt == 4:
                raise RuntimeError(f"Graph {method} request failed: {error}") from error
            delay = 2**attempt
            print(f"Graph request failed; retrying in {delay}s...", file=sys.stderr)
            time.sleep(delay)
            continue

        if response.status_code in (429, 500, 502, 503, 504):
            if attempt == 4:
                raise RuntimeError(
                    f"Graph {method} failed ({response.status_code}): {response.text}"
                )
            retry_after = response.headers.get("Retry-After")
            delay = int(retry_after) if retry_after is not None else 2**attempt
            print(f"Graph returned {response.status_code}; retrying in {delay}s...", file=sys.stderr)
            time.sleep(delay)
            continue

        if not response.ok:
            raise RuntimeError(f"Graph {method} failed ({response.status_code}): {response.text}")
        return response.json() if response.content else {}

    raise AssertionError("unreachable")


def folder_id_for_path(token: str, folder_path: str) -> str:
    """Resolve a slash-separated folder display-name path to a Graph ID."""
    url = f"{GRAPH_ROOT}/users/{quote(MAILBOX, safe='')}/mailFolders"
    parent_id: str | None = None
    for position, name in enumerate(part.strip() for part in folder_path.split("/")):
        if not name:
            raise ValueError(f"Invalid folder path: {folder_path!r}")

        # Graph accepts these stable names as folder IDs. This avoids relying
        # on translated display names such as "Sent Items" or "Gesendete Elemente".
        well_known_name = name.casefold()
        if position == 0 and well_known_name in WELL_KNOWN_FOLDERS:
            parent_id = well_known_name
            url = (
                f"{GRAPH_ROOT}/users/{quote(MAILBOX, safe='')}/mailFolders/"
                f"{well_known_name}/childFolders"
            )
            continue

        escaped_name = name.replace("'", "''")
        query = urlencode({"$select": "id,displayName", "$filter": f"displayName eq '{escaped_name}'"})
        response = graph_request(f"{url}?{query}", token)
        matches = response.get("value", [])
        if len(matches) != 1:
            location = "mailbox root" if parent_id is None else f"folder ID {parent_id}"
            raise RuntimeError(f"Expected one folder named {name!r} under {location}; found {len(matches)}.")
        folder_id = matches[0].get("id")
        if not isinstance(folder_id, str):
            raise RuntimeError(f"Graph returned a folder without a valid ID: {name!r}")
        parent_id = folder_id
        url = (
            f"{GRAPH_ROOT}/users/{quote(MAILBOX, safe='')}/mailFolders/"
            f"{quote(folder_id, safe='')}/childFolders"
        )
    assert parent_id is not None  # An empty path is rejected in the loop above.
    return parent_id


def access_token() -> str:
    """Obtain a delegated token and print device-login instructions."""
    app = msal.PublicClientApplication(
        CLIENT_ID, authority=f"https://login.microsoftonline.com/{TENANT_ID}"
    )
    flow = app.initiate_device_flow(scopes=SCOPES)
    if "user_code" not in flow:
        raise RuntimeError(f"Could not start device-code login: {flow}")
    print(flow["message"])
    result = app.acquire_token_by_device_flow(flow)
    if "access_token" not in result:
        raise RuntimeError(f"Login failed: {result.get('error_description', result)}")
    return result["access_token"]


def main() -> None:
    start = date.fromisoformat(START_DATE)
    end_exclusive = date.fromisoformat(END_DATE) + timedelta(days=1)
    if end_exclusive <= start:
        raise ValueError("END_DATE must not be before START_DATE")
    if not SOURCES:
        raise ValueError("SOURCES must contain at least one (folder, date field) pair")

    token = access_token()
    destination_id = folder_id_for_path(token, DESTINATION_FOLDER)

    # First collect every matching message. Do not change a source folder while
    # following its @odata.nextLink pages, otherwise Graph pagination can skip
    # messages after earlier page entries have been moved.
    messages_by_source: list[tuple[str, list[dict[str, Any]]]] = []
    for source_folder, date_field in SOURCES:
        if date_field not in {"sentDateTime", "receivedDateTime"}:
            raise ValueError(
                f"Date field for {source_folder!r} must be 'sentDateTime' or 'receivedDateTime'"
            )
        source_id = folder_id_for_path(token, source_folder)
        if source_id == destination_id:
            raise ValueError(f"Source folder {source_folder!r} resolves to the destination folder")

        date_filter = (
            f"{date_field} ge {start.isoformat()}T00:00:00Z and "
            f"{date_field} lt {end_exclusive.isoformat()}T00:00:00Z"
        )
        message_query = urlencode(
            {"$select": f"id,subject,{date_field}", "$filter": date_filter, "$top": "100"}
        )
        messages_url = (
            f"{GRAPH_ROOT}/users/{quote(MAILBOX, safe='')}/mailFolders/"
            f"{quote(source_id, safe='')}/messages?{message_query}"
        )
        messages: list[dict[str, Any]] = []
        while messages_url:
            page = graph_request(messages_url, token)
            for message in page.get("value", []):
                message_id = message.get("id")
                if not isinstance(message_id, str):
                    raise RuntimeError(f"Graph returned a message without a valid ID in {source_folder!r}")
                messages.append(
                    {
                        "id": message_id,
                        "subject": message.get("subject") or "(no subject)",
                        "date": message.get(date_field),
                    }
                )
            messages_url = page.get("@odata.nextLink")
        messages_by_source.append((source_folder, messages))
        print(f"Collected {len(messages)} message(s) from {source_folder!r}.")

    if DRY_RUN:
        for source_folder, messages in messages_by_source:
            for message in messages:
                print(f"{source_folder}: Would move: {message['date']}  {message['subject']}")
        total = sum(len(messages) for _, messages in messages_by_source)
        print(f"Total: {total} message(s) would be moved to {DESTINATION_FOLDER!r}.")
        return

    # Move only after all source folders have been completely enumerated.
    results: list[tuple[str, int, int]] = []  # (folder, selected, successfully moved)
    for source_folder, messages in messages_by_source:
        successful_moves = 0
        for message in messages:
            print(f"{source_folder}: Moving: {message['date']}  {message['subject']}")
            try:
                graph_request(
                    f"{GRAPH_ROOT}/users/{quote(MAILBOX, safe='')}/messages/"
                    f"{quote(message['id'], safe='')}/move",
                    token,
                    method="POST",
                    body={"destinationId": destination_id},
                )
            except RuntimeError as error:
                # Continue so the end-of-run check can identify all failures.
                print(f"{source_folder}: FAILED to move {message['id']}: {error}", file=sys.stderr)
            else:
                successful_moves += 1
        results.append((source_folder, len(messages), successful_moves))

    selected_total = sum(selected for _, selected, _ in results)
    moved_total = sum(moved for _, _, moved in results)
    for source_folder, selected, moved in results:
        status = "OK" if selected == moved else "MISMATCH"
        print(f"{source_folder!r}: {moved} of {selected} message(s) moved to {DESTINATION_FOLDER!r} [{status}].")
    print(f"Total: {moved_total} of {selected_total} message(s) moved.")
    if moved_total != selected_total:
        raise RuntimeError("Not all collected messages were moved; see the failure messages above.")


if __name__ == "__main__":
    main()
