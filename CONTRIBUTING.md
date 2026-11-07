# Contributing

Thanks for taking the time to contribute. This document is short on purpose.

## Commit messages

Format:

```
type(scope): imperative summary

Why the change was needed and why this approach was chosen.
Not a restatement of the diff.

Tested: what was run and what was observed.

Assisted-by: <tool> (<model>)
Closes #<issue>
```

- `type` is one of `feat`, `fix`, `refactor`, `docs`, `test`, `ci`, `chore`, `perf`
- the summary is imperative ("fix sorting", not "fixed sorting"), max 72 characters, no trailing period
- body wraps at 72 characters and states **why**, plus what was actually verified
- one logical change per commit; do not mix a refactor with a feature

A template is included in this repository:

```bash
git config commit.template .gitmessage
```

## AI-generated contributions

AI assistance is allowed. Two rules apply:

1. **A human is accountable.** Whoever pushes the change has read the whole diff, understands it, and can answer questions about it during review. Do not submit code you cannot explain.
2. **Disclose it.** When a meaningful part of the change was written by an assistant, add a trailer to every affected commit:

   ```
   Assisted-by: <tool> (<model>)
   ```

   and fill in the "How this was made" section of the pull request template.

Additional conventions:

- `Signed-off-by` is a human statement (Developer Certificate of Origin). An AI assistant must never add it; only the person pushing the commit may.
- `Co-authored-by` is only meaningful for a person whose email is linked to a GitHub account. Do not use it to credit a model.
- Say plainly in the commit body what was **not** tested. An honest gap is better than a false claim.

## Testing before you open a pull request

This is a desktop GUI program, so most changes need a manual check:

```bash
pip install -r requirements.txt   # or install customtkinter and pillow
python PnP2PDF.py
```

Build a PDF from `sample-project` (or your own cards folder) and confirm the output: page count, card order, cut lines, and — for double-sided cards — that fronts and backs line up.

State in the commit or PR what you ran. "Not tested on Windows/macOS" is a valid and useful sentence.

## Pull requests

- Keep them focused; one concern per pull request
- Reference the issue it resolves with `Closes #<number>`
- Screenshots of a changed layout are welcome

## Running the tests

```bash
python -m pip install -r requirements-dev.txt
python -m pytest
```

The suite stubs customtkinter, so no display is needed.
