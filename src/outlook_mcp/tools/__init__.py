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
