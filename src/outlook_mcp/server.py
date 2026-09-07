"""FastMCP server construction and tool registration."""

from __future__ import annotations

from mcp.server.fastmcp import FastMCP

from outlook_mcp.bridge import OutlookBridge
from outlook_mcp.tools import enabled_tool_docs, register_all
from outlook_mcp.ui import register_ui

INSTRUCTIONS_TEMPLATE = """\
This MCP server gives you access to Microsoft Outlook desktop on
Windows via COM automation. The tool groups registered in this session
are listed below; anything not listed is switched off and unavailable,
so do not offer it. In particular, mail tools are absent unless the
Mail category appears below.

Authentication piggybacks on whatever account Outlook is signed into —
no Microsoft Graph API, no Entra app registration, no OAuth tokens.

PREREQUISITE: classic Outlook (OUTLOOK.EXE) must be installed. The
server auto-launches Outlook on first call; the new "modern" Outlook
(olk.exe) is NOT supported.

Tool categories (all prefixed `outlook_`):
{tool_categories}

Most read tools accept response_format='markdown' (default) or 'json'.
Item references use Outlook EntryID strings; list tools return them on
every item — pass them back to detail/edit/delete tools.

All datetimes are in the USER'S LOCAL timezone with an explicit UTC
offset (e.g. 2026-06-10T16:33:22+05:00). Present them as-is — do NOT
convert to another timezone. outlook_whoami reports the timezone name
and current local time.
"""


def build_server() -> tuple[FastMCP, OutlookBridge]:
    """Construct the FastMCP instance, bridge, and wire all tools.

    The bridge is *not* started here — entrypoint.main() does that so
    the readiness wait happens after logging is configured.
    """
    instructions = INSTRUCTIONS_TEMPLATE.format(tool_categories=enabled_tool_docs())
    mcp = FastMCP("outlook_mcp", instructions=instructions)
    bridge = OutlookBridge()
    register_all(mcp, bridge)
    register_ui(mcp)
    return mcp, bridge
