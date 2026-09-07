"""Tool registration: each submodule exposes register(mcp, bridge)."""

import logging
import os

from . import (
    account,
    calendar,
    categories,
    contacts,
    folders,
    mail,
    ooo,
    rules,
    tasks,
)

logger = logging.getLogger("outlook_mcp.tools")

# Registration order is significant only for how tools appear to the host.
_GROUPS = {
    "mail": mail,
    "folders": folders,
    "calendar": calendar,
    "contacts": contacts,
    "tasks": tasks,
    "categories": categories,
    "rules": rules,
    "ooo": ooo,
    "account": account,
}


# One line per group, mirroring the tool names each module registers.
# Kept beside _GROUPS so the two cannot drift apart.
_GROUP_DOCS = {
    "mail": "Mail: list, search, get, send, reply, forward, move, delete, mark, save_attachments",
    "folders": "Folders: list_folders, create_folder",
    "calendar": "Calendar: list_events, get_event, create_event, update_event, delete_event, respond_event",
    "contacts": (
        "Contacts: list_contacts, search_contacts (incl. org directory/GAL), "
        "get_contact, resolve_name (name -> SMTP address)"
    ),
    "tasks": "Tasks: list_tasks, create_task, complete_task",
    "categories": "Categories: list_categories, set_category",
    "rules": "Rules: list_rules, toggle_rule  (modifies live mail rules \u2014 confirm first)",
    "ooo": "Out-of-Office: get_out_of_office",
    "account": "Account: whoami  (sanity check on the bound mailbox)",
}


def _enabled_groups() -> tuple[str, ...]:
    """Tool groups to register, from OUTLOOK_MCP_TOOLS (default: all).

    Set e.g. OUTLOOK_MCP_TOOLS=calendar,account to expose only the
    calendar tools, so they cannot collide with another mail server's
    tools in the same session.
    """
    raw = os.environ.get("OUTLOOK_MCP_TOOLS", "").strip()
    if not raw:
        return tuple(_GROUPS)
    wanted = {part.strip().lower() for part in raw.split(",") if part.strip()}
    unknown = wanted - set(_GROUPS)
    if unknown:
        raise ValueError(
            f"OUTLOOK_MCP_TOOLS names unknown tool groups: {sorted(unknown)}. "
            f"Valid groups: {sorted(_GROUPS)}"
        )
    return tuple(name for name in _GROUPS if name in wanted)


def register_all(mcp, bridge) -> None:
    names = _enabled_groups()
    logger.info("Registering tool groups: %s", ", ".join(names))
    for name in names:
        _GROUPS[name].register(mcp, bridge)


def enabled_tool_docs() -> str:
    """Bullet list of the tool groups actually registered, for the
    server instructions. Without this the host is told about mail and
    rules tools that OUTLOOK_MCP_TOOLS has switched off."""
    return "\n".join(f"  - {_GROUP_DOCS[name]}" for name in _enabled_groups())
