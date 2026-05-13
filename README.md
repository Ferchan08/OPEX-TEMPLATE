# OpEx Charter Auto-Fill

Automatically fills the OpEx Charter PowerPoint template from a JSON file.  
Every time `data.json` is updated and pushed, GitHub Actions runs the script and produces a ready-to-download `filled_charter.pptx`.

---

## Repository structure

```
├── template.pptx            ← Your PowerPoint template (do NOT rename)
├── data.json                ← JSON data from your Glean agent
├── fill_opex_template.py    ← Python script that does the filling
├── requirements.txt         ← Python dependencies
└── .github/
    └── workflows/
        └── fill_charter.yml ← GitHub Actions workflow
```

---

## How it works

1. Your **Glean agent** produces a JSON with project data.
2. You (or an automation) update `data.json` and push to this repo.
3. **GitHub Actions** runs `fill_opex_template.py` automatically.
4. The filled PPTX is available as a downloadable **artifact** in the Actions tab.

---

## JSON keys reference

| Key | Description | Shape in slide |
|-----|-------------|----------------|
| `project_name` | Project title | ProjectNameBox |
| `problem_statement` | Problem description | ProblemStatementBox |
| `project_leader` | Leader name | ProjectLeaderBox |
| `team_members` | Team member list | TeamMembersBox |
| `project_objectives` | Objectives | ProjectObjectivesBox |
| `project_timing` | Timing narrative | ProjectTimingBox |
| `project_sponsors` | Sponsor(s) | ProjectSponsorsBox |
| `opex_master` | OpEx master name | OpExMasterBox |
| `in_scope` | In-scope description | InScopeBox |
| `out_of_scope` | Out-of-scope description | OutOfScopeBox |
| `project_impact_3y` | 3-year impact narrative | ProjectImpactBox |
| `finance` | Finance contact | FinanceBox |
| `define` | Define phase date (YYYY-MM-DD) | Define |
| `measure` | Measure phase date (YYYY-MM-DD) | Measure |
| `analyze` | Analyze phase date (YYYY-MM-DD) | Analyze |
| `improve` | Improve phase date (YYYY-MM-DD) | Improve |
| `control` | Control phase date (YYYY-MM-DD) | Control |

---

## Run locally

```bash
# Install dependency
pip install lxml

# Fill the template
python fill_opex_template.py \
  --json data.json \
  --template template.pptx \
  --output filled_charter.pptx
```

---

## Download the filled PPTX from GitHub Actions

1. Go to the **Actions** tab in this repository.
2. Click the latest **Fill OpEx Charter** workflow run.
3. Scroll to **Artifacts** → click **filled-opex-charter** to download.

---

## Trigger manually

You can also trigger the workflow manually:

1. Go to **Actions** → **Fill OpEx Charter**.
2. Click **Run workflow** → **Run workflow**.
