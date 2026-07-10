# Installing the Three-Statement Financial Model skill

The skill builds a best-practice, self-validating multi-sheet three-statement model
workbook (Cover · Model · Checks/audit · Dashboard · Sensitivity) with
jurisdiction-aware tax/accounting for Egypt and the GCC. Pick whichever install
method fits how you run Claude.

**Runtime requirements** (all methods): `python3` with `openpyxl` to build; and
**either** LibreOffice headless (`soffice`) **or** the `formulas` pip package to
validate. Install the light option with `pip install openpyxl formulas`.

---

## Method 1 — Claude Code plugin (one-command, recommended)

This repository is a Claude Code plugin marketplace. In an interactive Claude Code
session:

```
/plugin marketplace add realElShamy/RegPy
/plugin install three-statement-model@regpy-marketplace
```

Then just ask for a financial model, or run `/three-statement-model`. Updates come
with `/plugin marketplace update regpy-marketplace`.

From the CLI (non-interactive):

```bash
claude plugin marketplace add realElShamy/RegPy
claude plugin install three-statement-model@regpy-marketplace
```

## Method 2 — Personal skill (available in every project)

Unzip the distributed bundle into your personal skills directory:

```bash
unzip three-statement-model-skill.zip -d ~/.claude/skills/
# → ~/.claude/skills/three-statement-model/SKILL.md
```

(Regenerate the zip yourself any time with `bash plugins/sync_plugin.sh --zip`,
which writes `dist/three-statement-model-skill.zip`.)

## Method 3 — Project skill (checked into a repo)

Copy the skill folder into a project so it loads automatically for anyone working
in that repo:

```bash
mkdir -p <your-project>/.claude/skills
cp -r .claude/skills/three-statement-model <your-project>/.claude/skills/
```

In this repository the skill already lives at
`.claude/skills/three-statement-model/`, so cloning RegPy and opening it in Claude
Code loads the skill with no further steps.

## Method 4 — Any other agent harness (portable)

The skill is self-contained and harness-agnostic. Copy the
`three-statement-model/` folder anywhere the agent can read files and run
`python3`, then point the agent at `SKILL.md` as its instructions. `SKILL.md`
phrases data-source discovery and file delivery as capability probes, so it adapts
to whatever tools the host exposes (it names Claude Code equivalents only as
examples).

---

## Verifying the install

Build and validate the bundled example (from inside the skill directory):

```bash
python3 scripts/build_workbook.py \
    --inputs assets/example_inputs.json \
    --jurisdiction assets/jurisdictions/eg.json \
    --output /tmp/example_model.xlsx

python3 scripts/validate_workbook.py /tmp/example_model.xlsx \
    --inputs assets/example_inputs.json \
    --jurisdiction assets/jurisdictions/eg.json
```

A clean install prints `OK: … Cover · Three Statement Model · Checks · Dashboard ·
Sensitivity` from the builder and `PASS: … FAIL: 0` from the validator.

## Layout

```
three-statement-model/
├── SKILL.md              # entry point: the orchestrating pipeline
├── references/           # TEMPLATE_SPEC, MODEL_ANATOMY, JURISDICTIONS,
│                         #   DATA_SOURCES, BEST_PRACTICES, DASHBOARD, SENSITIVITY
├── scripts/              # build_workbook / validate_workbook (full workbook),
│                         #   build_model / validate_model (core single sheet)
└── assets/               # inputs.schema.json, example_inputs.json,
                          #   jurisdictions/{eg,sa,ae,qa,kw,bh,om,intl}.json
```

The canonical source is `.claude/skills/three-statement-model/`. The
`plugins/three-statement-model/` copy is the plugin distribution artifact, kept in
sync by `plugins/sync_plugin.sh` — edit the canonical copy, then re-run the script.
