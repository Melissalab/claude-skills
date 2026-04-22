# Claude Skills

My personal library of skills for Claude. Each skill teaches Claude how to handle a specific kind of task consistently well — producing a meal plan document, drafting a proposal, running a CRM workflow, or whatever else I build over time. This repository acts as both my backup and the source of truth that Claude Code, Cowork, and Claude.ai all pull from.

## Available skills

### meal-plan

Creates personalized weekly meal plan documents tailored to a specific person, diet type, and grocery source. Modular and extensible: new diet types (diabetic, weight loss, muscle gain), new grocery sources (Kroger, Whole Foods, Amazon Fresh), and new person profiles can be added over time by creating additional reference files. The output is a polished Word document with a cover page, quick-start guide, shopping list organized by store section, seven-day meal table with calories and sugar per meal, recipes under twenty minutes each, and an exhaustive grocery reference.

## How to use these skills

### In Claude Code

Open Claude Code in your terminal and run the following commands once.

```
/plugin marketplace add <your-github-username>/claude-skills
/plugin install meal-plan@melissa-skills
```

After this, the skill activates automatically when you ask Claude Code to make a meal plan.

### In Cowork

Open the Cowork desktop app. Click Customize in the left sidebar, then click the plus button to add a new plugin source. Paste in `<your-github-username>/claude-skills` as the marketplace URL. The meal-plan plugin will appear, and you can install it with one click. After installation the skill activates automatically when relevant.

### In Claude.ai (web chat)

Claude.ai does not yet read directly from GitHub plugin marketplaces. To use a skill here, you need to upload the packaged version manually. Navigate to Customize then Skills in Claude.ai settings, click Upload, and select the `.skill` file for the skill you want to install. The packaged `.skill` files for each skill are available in this repository as release attachments.

## Adding new skills

Each new skill becomes a new folder under `plugins/`, with its own `.claude-plugin/plugin.json` manifest and a `skills/` subfolder containing the actual skill. After adding a new plugin folder, update `.claude-plugin/marketplace.json` at the repo root to include the new plugin in the list.
