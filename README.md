# Outlook MCP Server

Local MCP server for Microsoft Outlook Classic on Windows, driven through COM (pywin32).
It lets an AI assistant read, triage and organise a mailbox while keeping every outbound
action as a draft for human review.

Fork lineage: [marlonluo2018/outlook-mcp-server](https://github.com/marlonluo2018/outlook-mcp-server)
→ [MagoDev/outlook-mcp](https://github.com/MagoDev/outlook-mcp) (draft-first guardrails) → this repository.

## Safety model

- **No sending.** Replies, new messages and forwards are saved as drafts in the Drafts folder.
  No tool sends mail, responds to meetings or creates calendar invitations.
- **Recoverable changes.** Deleting a message or folder moves it to Deleted Items. Saving files
  to disk never overwrites an existing file.
- **Audit trail.** Every write (draft, move, category, flag, task, saved attachment) is appended
  to `~/.outlook-mcp/audit.log` on the local machine.
- **Local only.** The server talks to the Outlook client already running on the machine. What the
  AI client does with the data it reads depends on that client, not on this server.

## Requirements

- Windows 10 or later
- Microsoft Outlook Classic (desktop), running and signed in
- Python 3.8 or later

## Installation

```bash
git clone https://github.com/Robvred/outlook-mcp.git
cd outlook-mcp
pip install -e .
```

An editable install (`-e`) means a later `git pull` is picked up after restarting the MCP client.

MCP client configuration (for example Claude Desktop):

```json
{
  "mcpServers": {
    "outlook": {
      "command": "python",
      "args": ["-m", "outlook_mcp_server"]
    }
  }
}
```

After updating the code, fully quit and restart the MCP client so the server is reloaded.

Quick checks:

```bash
python -c "import outlook_mcp_server; print('module ok')"
python -c "import win32com.client as w; w.Dispatch('Outlook.Application'); print('outlook ok')"
```

## Tools

### Read and search

| Tool | Purpose |
|---|---|
| `list_recent_emails_tool` | Load the last N days of a folder (max 30) into the cache |
| `load_emails_by_folder_tool` | Load a folder by period or count |
| `search_email_by_subject_tool`, `..._by_sender_name_tool`, `..._by_recipient_name_tool`, `..._by_body_tool` | Search and load results into the cache |
| `view_email_cache_tool` | Page through the cache (5 per page), with categories and flag status |
| `get_email_by_number_tool` | Read one message, optionally with thread and attachments |
| `clear_email_cache_tool` | Clear the cache |
| `mailbox_stats_tool` | Read-only statistics over up to 730 days (volume, senders, conversation topics), without using the cache |

### Drafts

| Tool | Purpose |
|---|---|
| `reply_to_email_by_number_tool` | Reply-all draft using Outlook's native reply: HTML, signature, quoted thread, threading and resolved recipients. Options: replace To/Cc (an empty Cc clears it), add Cc, importance. Returns the final recipient list |
| `forward_email_by_number_tool` | Forward draft using Outlook's native forward: original message and attachments included, explicit To/Cc, optional text on top. Use it to bring someone into a topic with its context |
| `compose_email_tool` | New draft in HTML with signature |
| `batch_forward_email_tool` | Forward drafts to a recipient list |

### Organise

| Tool | Purpose |
|---|---|
| `get_folder_list_tool`, `create_folder_tool`, `move_folder_tool`, `remove_folder_tool` | Folder management |
| `move_email_tool`, `delete_email_by_number_tool` | Move or delete one message |
| `move_emails_by_filter_tool` | Bulk move by period, subject, sender, excluding categories. Dry run by default, time-boxed batches |
| `ensure_categories_tool`, `set_email_categories_tool` | Create categories, tag messages (add, replace, remove, exclusive group) |

### Follow-up

| Tool | Purpose |
|---|---|
| `flag_email_tool` | Flag, complete or clear, with due date and reminder |
| `create_task_tool` | Outlook task with due date, reminder, categories and the source message attached |
| `list_todo_tool` | Open tasks and flagged messages sorted by due date |
| `complete_todo_tool` | Mark a task or flag as complete |

### Files

| Tool | Purpose |
|---|---|
| `save_attachments_tool` | Save real attachments to a local folder (no overwrite) and create `.url` shortcuts for online file links |
| `list_local_directory_tool` | List a local folder tree to choose where to file |

## Operating notes

- **One shared cache.** Every load or search replaces the cache; message numbers refer to the last
  load only. Do not run cache-loading calls in parallel.
- **Time limit.** MCP clients usually time out after about 60 s. Long operations are time-boxed;
  repeat the call until it reports nothing remaining.
- **Matching.** Subject and sender filters are case-insensitive substring matches. Check the dry-run
  sample before moving.
- **Dates.** Due dates and reminders are given in local time (`YYYY-MM-DD` or `YYYY-MM-DD HH:MM`).
- **Categories.** On Exchange, categories created through COM may disappear from the master list
  after a restart; messages keep their tags. Run `ensure_categories_tool` again or create them once
  from the Outlook UI.

## Bulk export (optional)

`outlook-export` exports a mailbox incrementally to `.msg` files with a SQLite full-text index:

```bash
outlook-export --archive <folder> --db <file.sqlite> --dry-run
```

Exports contain mailbox data: keep them out of this repository (see `.gitignore`).

## Contributing

This repository is public. Read [CLAUDE.md](CLAUDE.md) first: no real organisations, people,
addresses, folder paths, projects, figures or mailbox content in code, examples, commits or pull
requests. Use the neutral placeholders listed there.

Further reading: [docs/](docs/) (COM implementation notes, performance, Graph API comparison).

## License

MIT, see the upstream project.
