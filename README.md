# MBE Updates

Standalone Python GUI consoles for Mail Boxes Etc. (MBE) logistics operations
(customs tracking, XML declaration generation, invoice splitting, and customer
cargo notifications). Each console is a single self-contained `.pyw` file that
users run on Windows. This repository also hosts the JSON `manifests/` that the
consoles poll for self-updates and the canonical (public) copies of the console
`.pyw` files that updates are downloaded from.

## Self-update + local-data preservation (the update contract)

Several consoles are **self-modifying**: they store live, user-entered data
*inside their own `.pyw` file* in a clearly marked data block. This is how the
apps work without any external database or Excel file.

Because that same `.pyw` file is also the thing the self-updater downloads and
replaces from GitHub, the update path is designed around one rule:

> **Splice local data in, never overwrite user entries, never publish customer
> data.**

This is enforced by `_download_and_apply_update` in each console. On update it:

1. Downloads the fresh console version from GitHub (raw URL in the manifest).
2. Re-reads the **local** `.pyw` currently running on the user's machine.
3. Extracts the local data block(s) from that local copy.
4. Splices those local blocks into the freshly downloaded text (replacing the
   empty block that ships in the public copy).
5. Atomically swaps the file into place with `os.replace`.

Two hard requirements follow from this:

- **Public copies ship empty.** The copies of the console `.pyw` files in this
  repository (and therefore on GitHub) MUST keep their data blocks empty so no
  customer information is ever published. A populated data block committed here
  would leak customer data *and* would be spliced out on the next user update
  anyway.
- **The marker names are part of the update contract.** The updater and the
  in-app save/load code both locate the data block by its exact marker
  (variable name or comment marker). If a marker is renamed, moved, or removed,
  the splice/save/load logic that references it MUST be updated in lockstep — or
  every user's entries will be silently wiped on the next update. Do not rename
  a marker as a drive-by change.

### Per-console preserved data blocks

| Console (`consoles/…`) | Preserved block(s) | Empty public form | Preserved by |
| --- | --- | --- | --- |
| `Packages at Customs Console.pyw` | `CUSTOMS_DATA` (list) | `CUSTOMS_DATA = []` | `_splice_block` (matched by `_DATA_BLOCK_PATTERN`) |
| `XML Declaration Console.pyw` | `BUILTIN_TIN_NUMBERS` (dict) **and** `BUILTIN_CODES` (list) | `BUILTIN_TIN_NUMBERS = {}` / `BUILTIN_CODES = []` | `_splice_block` (matched by `_TIN_DATA_PATTERN` / `_CODES_DATA_PATTERN`) |
| `notify customers of ocean cargo on hand.pyw` | `EMAIL_LOOKUP` (dict), delimited by `CLIENT_MARKER_START` / `CLIENT_MARKER_END` | `EMAIL_LOOKUP = {}` | marker splice inside `_download_and_apply_update` |
| `Split Factura by CBY (Mass Print).pyw` | *none* — plain overwrite | n/a | no data block; see below |

`Split Factura by CBY (Mass Print).pyw` intentionally has **no** user-edited
embedded data block: it processes selected PDF files and holds all state in
memory only, never writing user data back into its `.pyw`. Its
`_download_and_apply_update` therefore does a plain full-file overwrite with no
splice. If a persisted/user-editable data block is ever added to that console,
give it the same splice-preservation treatment as the others.

### Merge-on-save (concurrent edits)

For the consoles whose data blocks are edited in-app, saving does more than
overwrite the block wholesale:

- **XML Declaration Console** (`BUILTIN_TIN_NUMBERS`, `BUILTIN_CODES`) and
  **Ocean Cargo notification** (`EMAIL_LOOKUP`) re-read the on-disk data block
  at save time and **merge** it with the in-memory edits before writing. This
  means if two users on shared storage (e.g. Dropbox) edit different entries,
  one user's save will not clobber the other's — entries added on disk that
  aren't in memory are kept.
- **Packages at Customs Console** (`CUSTOMS_DATA`) rewrites its block from the
  in-memory rows on save. Its live data is still fully preserved across
  self-updates via the splice described above.

In all cases the self-update splice is the safety net that carries live local
data across a code replacement.

## Repository layout

- `consoles/` — the console applications (`.pyw`). Public copies ship with
  **empty** data blocks (see above).
- `manifests/` — one JSON file per console with `version`, `changelog`, and the
  raw GitHub `url` the updater downloads. Bump `version` here to roll out an
  update.
