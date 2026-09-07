# Local fork status — outlook-classic-mcp

Last updated: 2026-09-07. Branch `local/calendar-only` in `C:\Code\outlook-classic-mcp`.
Setup: classic desktop Outlook, local PST file only, no online/Exchange account.

**Everything below is committed, live and verified** as of the 2026-09-07 restart of Outlook and the Claude UI.

## Local commits on top of upstream

| Commit | What |
|---|---|
| `bb687b4` | Gate tool groups behind `OUTLOOK_MCP_TOOLS` |
| `f39c459` | Support creating all-day events (Outlook's exclusive-end semantics handled in code) |
| `e6ca222` | Generate server instructions from the enabled tool groups |
| `fe1c2b9` | Fix locale-dependent date filter in `list_events` |
| `345d9fa` | Fix timezone shift when writing dates to Outlook |
| `35c59ca` | Surface and set categories on calendar events |

Enabled groups: `calendar`, `categories`. Email is deliberately off (the PST has no mail, and it would collide with the Thunderbird MCP). Address book is not enabled yet and is the obvious next piece of work.

## Two locale/timezone traps, both fixed — read before touching date code

**1. Date filter (`fe1c2b9`).** The old Jet restriction embedded dates as `%m/%d/%Y`, and Outlook parses that literal using the Windows locale, so `09/07/2026` was read as 9 July on this de-DE machine. It failed silently and *widened* the window, and only bit when the day of month was <= 12 (above that the day-first reading is invalid, Outlook falls back to month-first, and the query happened to be right).

Now: DASL against `urn:schemas:calendar:dtstart` with ISO literals, which are not locale-parsed, plus a result-loop filter on the parsed start so the returned set is exact whatever `Restrict` did. `Restrict` only bounds how far recurrence expansion runs, so an unusable filter is downgraded to a full scan instead of failing the call. The ascending `items.Sort("[Start]")` is what makes the loop's `break` terminate on an open-ended series with `IncludeRecurrences`.

Verified: 07–09 Sep, 20–28 Sep (day > 12) and 30 Nov – 06 Dec (month and DST boundary).
**The old "swap day and month" workaround is now wrong.** Never reintroduce it.

**2. Timezone on write (`345d9fa`).** pywin32 converts whatever datetime it is handed to UTC and gives Outlook that clock, which Outlook stores as local wall time. Passing the naive local value from `from_iso` therefore stored the UTC representation, so every event landed one UTC offset early (2 h CEST, 1 h CET). All-day was worse: midnight shifted to 22:00 the previous day and `AllDayEvent = True` re-snapped the span outwards, turning one day into two starting a day early.

Now: `to_com` in `utils/formatting.py`, the exact inverse of `to_iso`. It labels the intended wall clock as UTC so pywin32's conversion is a no-op. A relabel, not arithmetic, so there is no DST edge case. Applied to appointment `Start`/`End` on create and update, `PatternEndDate`, and task `DueDate`/`ReminderTime`. Round trip covered in `tests/test_formatting.py`.

Verified after restart: a 14:00 request stored 14:00, and all-day `2026-10-16`..`2026-10-17` stored exactly 16 Oct.

**The rule both bugs share:** never hand Outlook a date whose interpretation depends on the machine's locale or timezone. Use ISO literals in DASL, and `to_com` for every COM date assignment.

### Do not "repair" old events

Events created through the MCP before `345d9fa` went live were written with the offset. The session that created them had already spotted the bug and pre-compensated the times, so **they are correct in the calendar now**. Do not shift them.

## Categories

Working end to end. The `categories` group (`outlook_list_categories`, `outlook_set_category`) is upstream and now enabled; `35c59ca` added the calendar-side pieces: `list_events` returns categories per event, and `create_event` / `update_event` take a `categories` argument.

Separator: **resolved, comma is fine.** Outlook here accepts both `,` and `;` on input and normalises to `", "` on output — `"Privat;Uni"` and `"Uni,Familie"` both came back joined with a comma and a space. The earlier worry that a German locale would need `;` was wrong. `set_category`'s "comma-separated" docstring is accurate; leave it. The one real ambiguity is a category name containing a comma, and none of the 14 names in this profile do.

The client code passes the string through verbatim rather than joining a list, which is why the input separator does not matter.

## Open items

1. **Inclusive vs exclusive end.** `list_events` treats `end` as inclusive, so 07–09 Sep returns three days and an all-day event starting at midnight on the 9th is included. `create_event` with `all_day` documents its end as exclusive. Same server, two conventions for the same field.
2. **Events overlapping the window but starting before it are missed.** Filtering is on start time only, so a vacation block from 5 to 10 September does not show when querying the 7th to the 9th. Predates all local patches. A fix means restricting on overlap (`dtstart < end AND dtend > start`), which gives up the `break`-based early termination and needs another guard against open-ended recurring series.
3. **`limit` (default 50) truncates silently.** No flag in the response says results were cut, so a busy month looks complete.
4. **`list_mails` still has the locale date bug** on `[ReceivedTime]`. Latent, since the email group is off. Fix before ever enabling email.
5. **Address book group** not enabled yet.
6. **`response_format` had no visible effect** on `list_events` through the Claude desktop bridge: asking for markdown still returned the structured JSON. `list_events` and `get_event` wrap their output in `ui_result`, which attaches structured content, and that client appears to render the structured content instead of the markdown text. Not investigated, and it may look different in Claude Code. Not obviously a server bug.

## Known non-bugs

`outlook_whoami` returns `current_user: "Unknown"` with an empty accounts list, and most events report `organizer: "Unknown"`. Expected for a PST-only profile, not a connection failure.

If both `whoami` and `list_events` fail with `AttributeError` on `Outlook.Application.Session` / `GetNamespace.GetDefaultFolder`, Outlook is not running, so there is no MAPI session. Start Outlook rather than retrying.

## Development

The venv at `.venv` was created by `install.bat` with the package installed editable (`-e .`), so code edits take effect on the next server start with no reinstall. The server is `.venv\Scripts\python.exe -m outlook_mcp` over stdio.

Tests: `uv run pytest` on Windows. They cannot run in the Linux VM that the Claude desktop bridge uses, which has neither pytest nor pydantic.

**A code change is not live until the MCP server process restarts**, which means restarting the client that spawned it. Every past round of confusion in this project came from testing against a server still running the old code. Verify a fix by creating a throwaway event, reading it back, and deleting it, rather than assuming the restart took.

Committing from the Claude desktop bridge needs `GIT_AUTHOR_NAME`/`GIT_AUTHOR_EMAIL` set inline, since git has no identity in that Linux VM. Committing from Windows does not.

## Continuing this work in Claude Code on Windows

Register the server once, from `cmd`:

```
claude mcp add outlook --scope user --env OUTLOOK_MCP_TOOLS=calendar,categories -- C:\Code\outlook-classic-mcp\.venv\Scripts\python.exe -m outlook_mcp
```

`--scope user` makes it available in every project rather than only the one it was added from. Check the `OUTLOOK_MCP_TOOLS` value against whatever the Claude Desktop config currently uses (`%APPDATA%\Claude\claude_desktop_config.json`) so the two clients expose the same tools. Verify with `claude mcp list`, and `/mcp` inside a session.

Note that `scripts/install_to_clients.py` can register Claude Code too, but it passes neither `--scope` nor `--env`, so it would land in project scope with all tool groups enabled. Prefer the explicit command above.

Claude Desktop and Claude Code each spawn their own server process; both attach to the same running Outlook over COM.

Tool approval in Claude Code is separate from the Claude Desktop connector settings: MCP tools are matched by full name, e.g. `mcp__outlook__outlook_create_event`, in the `permissions` block of `.claude\settings.json`, or via `/permissions`.
