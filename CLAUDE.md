# outlook-classic-mcp — local fork

Local fork of a third-party MCP server, on branch `local/calendar-only`, run against classic desktop Outlook with a local PST file (no Exchange, no online account).

**Read `LOCAL-STATUS.md` before changing anything.** It carries the state of the local patches, what has been verified against the live calendar, and the open items.

Standing rules that the two bugs fixed so far both came down to:

- **Never hand Outlook a date whose meaning depends on the machine's locale.** Filter with DASL and ISO literals, not the Jet `[Start] >= '...'` syntax. A US-formatted literal is silently misread on this de-DE machine, and only when the day of month is <= 12, which makes it look intermittent.
- **Never assign a naive datetime to a COM date property.** Use `to_com` from `utils/formatting.py`. pywin32 converts what it is given to UTC and Outlook stores that clock as local wall time, so a naive value lands one UTC offset early.
- **A code change is not live until the MCP server process restarts**, which means restarting the client that spawned it. Verify with a throwaway event: create, read back, delete.

Tests are `uv run --with pytest pytest`, Windows only. The venv is an editable install, so no reinstall is needed after an edit.
