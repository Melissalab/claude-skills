---
name: meal-plan
description: Create personalized weekly meal plan documents tailored to a specific person, diet type (diabetic, weight loss, muscle gain, heart healthy, and more), and grocery source (Kroger, Whole Foods, Amazon Fresh, Costco, and more). Use this skill whenever the user asks for a meal plan, weekly menu, shopping list with recipes, diet guide, or nutrition document built for a specific person (a friend, family member, client, or themselves). The output is a polished Word document in a warm friend-to-friend voice covering a cover page, quick-start guide, store-organized shopping list, 7-day meal table with calories and sugar per meal, under-20-minute recipes, and an exhaustive grocery reference. The skill is modular and extensible. New diet types, new grocery sources, and new person profiles can be added over time by creating reference files without changing the build logic.
---

# Meal Plan Builder

This skill produces personalized weekly meal plan documents. The underlying architecture is deliberately modular: the skill combines three independent pieces (a person profile, a diet type, and a grocery source) with universal tone rules and a build script to produce one custom document. This modularity is the whole point. It means a new diet type can be added as one new file, a new grocery source can be added as one new file, and person profiles accumulate over time as a personal library, all without touching the core build logic.

## When to use this skill

Trigger on any request for a meal plan, weekly menu, shopping list with recipes, diet guide, or nutrition document tailored to a specific person. Common phrasings include asking Claude to make a meal plan for a friend, a spouse, a parent, a client, or the user themselves. The tell is specificity: there is usually a named person on the receiving end and often a named diet type or health goal. Abstract questions about nutrition do not need this skill; conversational answers are better for those.

## Architecture

The skill directory is organized into the following shape.

The SKILL.md file at the top (this file) describes the workflow for combining variants and pointers to where each variant lives. The references folder contains all the modular content organized into three subfolders. The diet-types subfolder holds one markdown file per nutritional approach (diabetic, weight loss, muscle gain, etc.), each capturing the calorie-targeting logic, key habits, and recipe library specific to that diet. The grocery-sources subfolder holds one markdown file per store (Kroger, Whole Foods, Amazon Fresh, Costco, etc.), each containing the product catalog organized by section with brand names actually available at that store. The profiles subfolder holds one markdown file per real person the user has built plans for, acting as an accumulating library so later plans can reference earlier context. A tone-and-style.md file sits at the references root because tone applies universally. The scripts folder contains the docx build script which is a pure assembly engine that reads configuration and produces the file.

## The workflow

When a user asks for a meal plan, follow this sequence.

First, identify the person. Check the profiles subfolder to see if this person already has a profile file. If they do, load it and confirm with the user whether their situation has changed since the last plan (new weight, new goals, new allergies). If they do not have a profile yet, walk through the intake questions in `references/profiles/_template.md` to capture the necessary data. Save the completed profile as a new markdown file in the profiles folder so future plans can reuse it.

Second, identify the diet type. Ask the user which nutritional approach fits (or infer from context if they said "my diabetic friend"). Load the matching file from `references/diet-types/`. If no existing file matches the person's situation (for example, they want a very specific kidney-disease-friendly plan that does not exist yet), create a new diet-type file by copying the structure of an existing one and adapting the nutritional rules, tips, and recipe library.

Third, identify the grocery source. Ask where the person shops, or default to the store in their profile if already known. Load the matching file from `references/grocery-sources/`. If the store does not exist yet in the library, you can either create a new grocery-source file by adapting an existing one, or substitute the closest available (Kroger works as a reasonable stand-in for most mainstream US supermarkets if no exact match exists).

Fourth, read `references/tone-and-style.md` to load the universal voice rules. This is essential for getting the prose right on the first pass.

Fifth, build the document by running the template script at `scripts/build_meal_plan.js`. Customize the CONFIG block and content sections by combining the profile data with the diet-type rules and the grocery-source product catalog. The script handles all the visual styling automatically.

Sixth, validate the output docx and present it to the user with the `present_files` tool.

## Document structure (applies to all variants)

Regardless of which diet type or grocery source, every meal plan document follows the same six-part structure. The cover page introduces the document with a warm title, subtitle, and intro paragraph, plus a "What's Inside" list of the other five sections. The second part is titled "A Few Things Worth Knowing" and contains the diet-specific tips in prose paragraphs with bolded lead-ins, plus a "How to Use This" explanation, plus a subtle italic caveat about individual reactions and checking with a doctor. The third part is the Weekly Shopping List, organized by store section with H2 headers and bullet items in "Item: quantity" format. The fourth part is the Weekly Meal Plan table with seven rows and five columns showing meal descriptions and calorie plus sugar counts. The fifth part is the Recipes section with breakfast, lunch, and dinner recipes each showing prep time, cook time, total time, and calorie estimate. The sixth part is the Full Grocery Reference, which is the exhaustive catalog of diet-appropriate options pulled from the grocery-source file.

## Visual styling (universal)

All meal plan documents use the same visual system. The font throughout is Verdana. Text is always black. The single accent color is Swiss Coffee (#E8E1D3), applied only behind header elements (H1 section headers, H2 subsection headers, and table header rows). Body table rows have no shading (no alternating zebra stripes). Em dashes and en dashes never appear anywhere in the document; they are replaced with colons for label-to-value relationships, commas for brief asides, periods for joining complete thoughts, and the word "to" for number ranges. Page size is US Letter with 0.75 inch margins.

The build script at `scripts/build_meal_plan.js` encodes all of this automatically. Do not change the styling layer; only change the content variables and sections.

## Adding a new diet type

To add a new diet type (for example, heart-healthy or low-FODMAP), create a new file at `references/diet-types/<name>.md` following the structure of `diabetic.md`. The file should contain four parts. First, the calorie and macronutrient targeting logic (what formula to use, what targets to aim for, what portion guidance scales with the target). Second, the "A Few Things Worth Knowing" content, which is four to six prose paragraphs with bolded lead-ins capturing the habits that matter most for this diet. Third, the seven-day meal structure template (what typical breakfasts, lunches, and dinners look like). Fourth, the full recipe library with prep times, cook times, total times, and calorie estimates.

## Adding a new grocery source

To add a new grocery source (for example, Trader Joe's or Aldi), create a new file at `references/grocery-sources/<name>.md` following the structure of `kroger.md`. The file should contain the product catalog organized by store section (produce, meat, dairy, frozen, pantry, etc.), each section opening with weekly quantity guidance and containing a "good to buy" list with real brand names and a "skip" list explaining what to avoid. Brands listed should actually be available at that store; substitute accordingly.

## Adding a new person profile

When the user asks for a plan for a new person, walk through the profile template at `references/profiles/_template.md` and save the completed profile as a new file named for the person (for example, `chad.md` or `lisa.md`). Profiles should capture everything known about the person that affects meal planning: name, age, sex, height, weight, activity level, occupation or lifestyle context, goals (maintain, lose, gain, manage specific condition), allergies and strong dislikes, equipment available (air fryer, stovetop, slow cooker, etc.), time budget for cooking, and any relevant notes from previous plans.

## Files in this skill

The file `references/tone-and-style.md` contains the universal voice rules. Read it early in every build, because retrofitting tone is harder than getting it right from the start.

The file `references/diet-types/diabetic.md` is the first diet-type variant, capturing everything specific to planning for Type 2 diabetes and A1C management.

The file `references/grocery-sources/kroger.md` is the first grocery-source variant, containing the full Kroger product catalog organized by store section.

The file `references/profiles/chad.md` is the first person profile, capturing the details for Chad (6 foot construction worker, 175 to 180 lbs, high A1C, Kroger shopper).

The file `references/profiles/_template.md` is the blank intake template for adding new profiles.

The file `scripts/build_meal_plan.js` is the document build engine. It reads configuration variables at the top and produces the final docx.
