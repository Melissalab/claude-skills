/**
 * Kroger Diabetic Meal Plan Builder
 *
 * This script generates a complete weekly meal plan document as a Word .docx file.
 *
 * HOW TO USE:
 *   1. Edit the CONFIG section below with the person's details
 *   2. Edit the CONTENT section if you need to change meals or add notes
 *   3. Run: node build_meal_plan.js
 *   4. Output goes to /mnt/user-data/outputs/meal_plan.docx
 *
 * Requires: npm install -g docx
 */

const fs = require('fs');
const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  AlignmentType, LevelFormat, HeadingLevel, BorderStyle, WidthType, ShadingType,
  PageOrientation, PageBreak
} = require('docx');

// ============================================================================
// CONFIG: Edit these variables for the specific person
// ============================================================================

const CONFIG = {
  // The person's first name, used in the document title
  personName: "Chad",

  // Warm one-liner that appears on the cover page under the title
  subtitle: "A full week of easy meals, one grocery run, nothing complicated.",

  // Daily calorie target. Calculate this from BMR * activity multiplier
  // minus 150-250 for A1C management. See SKILL.md Step 2.
  dailyCalorieTarget: 2490,

  // Output file path (defaults to /mnt/user-data/outputs for presenting to user)
  outputPath: "/mnt/user-data/outputs/meal_plan.docx",
};

// ============================================================================
// VISUAL STYLE: Do not change unless user explicitly asks
// These are the established style rules for this skill
// ============================================================================

const FONT = "Verdana";                    // Heavy, wide sans-serif, universally available
const SWISS_COFFEE = "E8E1D3";             // The single accent color, used for header bands only
const TEXT_COLOR = "000000";               // All text black, always

// ============================================================================
// HELPER FUNCTIONS: Used throughout content assembly
// ============================================================================

const border = { style: BorderStyle.SINGLE, size: 6, color: "BFBFBF" };
const borders = { top: border, bottom: border, left: border, right: border };

/** Plain paragraph of body text. */
const para = (text, opts = {}) => new Paragraph({
  children: [new TextRun({ text, font: FONT, color: TEXT_COLOR, ...opts })],
  spacing: { after: 80 },
});

/** Bullet point. Uses proper docx numbering config, never unicode bullets. */
const bullet = (text, opts = {}) => new Paragraph({
  numbering: { reference: "bullets", level: 0 },
  children: [new TextRun({ text, font: FONT, color: TEXT_COLOR, ...opts })],
  spacing: { after: 40 }
});

/** Section heading. Gets a full-width Swiss Coffee band behind it. */
const h1 = (text) => new Paragraph({
  heading: HeadingLevel.HEADING_1,
  children: [new TextRun({ text, font: FONT })],
  spacing: { before: 240, after: 160 },
  shading: { fill: SWISS_COFFEE, type: ShadingType.CLEAR }
});

/** Subsection heading. Same Swiss Coffee band as h1, smaller size. */
const h2 = (text) => new Paragraph({
  heading: HeadingLevel.HEADING_2,
  children: [new TextRun({ text, font: FONT })],
  spacing: { before: 200, after: 120 },
  shading: { fill: SWISS_COFFEE, type: ShadingType.CLEAR }
});

/** Minor heading. No shading. */
const h3 = (text) => new Paragraph({
  heading: HeadingLevel.HEADING_3,
  children: [new TextRun({ text, font: FONT })],
  spacing: { before: 160, after: 80 }
});

/** Italic note line, used for quantity guidance and recipe timing. */
const note = (text) => new Paragraph({
  children: [new TextRun({ text, italics: true, color: TEXT_COLOR, font: FONT, size: 20 })],
  spacing: { after: 120 }
});

const blank = () => new Paragraph({ children: [new TextRun("")] });

/** Build a table cell. Header cells get Swiss Coffee; body cells are plain. */
const cell = (text, opts = {}) => new TableCell({
  borders,
  width: { size: opts.width || 1000, type: WidthType.DXA },
  shading: opts.header ? { fill: SWISS_COFFEE, type: ShadingType.CLEAR } : undefined,
  margins: { top: 80, bottom: 80, left: 120, right: 120 },
  children: [new Paragraph({
    children: [new TextRun({
      text,
      bold: opts.header || opts.bold,
      color: TEXT_COLOR,
      font: FONT,
      size: 20
    })]
  })]
});

// ============================================================================
// CONTENT: The actual document contents
// Edit meal descriptions, shopping quantities, and notes here
// ============================================================================

const children = [];

// ---------------------------------------------------------------------------
// COVER PAGE
// ---------------------------------------------------------------------------

children.push(new Paragraph({
  alignment: AlignmentType.CENTER,
  children: [new TextRun({ text: "", size: 40 })],
  spacing: { after: 400 }
}));

children.push(new Paragraph({
  alignment: AlignmentType.CENTER,
  children: [new TextRun({
    text: `${CONFIG.personName}'s Weekly Meal Plan`,
    bold: true, size: 52, font: FONT
  })],
  spacing: { after: 160 }
}));

children.push(new Paragraph({
  alignment: AlignmentType.CENTER,
  children: [new TextRun({
    text: CONFIG.subtitle,
    italics: true, size: 26, font: FONT
  })],
  spacing: { after: 480 }
}));

// Intro paragraph - rewrite this for each person to fit their situation
children.push(new Paragraph({
  alignment: AlignmentType.CENTER,
  children: [new TextRun({
    text: "Here's a full week of meals put together with you in mind. One grocery run at Kroger, everything under 20 minutes to make, and enough variety that you won't get bored by Thursday. Nothing here needs more than a handful of ingredients, and the whole plan is built around foods you already like: wings, sausage, steak, cheese, and eggs, just pulled together in a way that keeps your blood sugar steady without making food feel like a chore.",
    size: 24, font: FONT
  })],
  spacing: { after: 400, line: 360 }
}));

// "What's Inside" table of contents
children.push(new Paragraph({
  alignment: AlignmentType.CENTER,
  children: [new TextRun({ text: "What's Inside", bold: true, size: 28, font: FONT })],
  shading: { fill: SWISS_COFFEE, type: ShadingType.CLEAR },
  spacing: { before: 240, after: 200 }
}));

const tocItems = [
  ["A Few Things Worth Knowing", "Small habits that make a big difference. Read this first."],
  ["1. Weekly Kroger Shopping List", "Everything you need for the week, organized by store section."],
  ["2. Weekly Meal Plan", "Seven days of breakfast, lunch, and dinner with calories and sugar."],
  ["3. Recipes", "Step-by-step for every meal. All under 20 minutes."],
  ["4. Full Kroger Reference", "Alternatives and swaps whenever you want to mix things up."],
];

tocItems.forEach(([title, desc]) => {
  children.push(new Paragraph({
    children: [
      new TextRun({ text: title + ": ", bold: true, size: 22, font: FONT }),
      new TextRun({ text: desc, size: 22, font: FONT })
    ],
    spacing: { after: 140 },
    indent: { left: 360 }
  }));
});

children.push(new Paragraph({ children: [new PageBreak()] }));

// ---------------------------------------------------------------------------
// PAGE 2: A FEW THINGS WORTH KNOWING
// ---------------------------------------------------------------------------

children.push(h1("A Few Things Worth Knowing"));

children.push(new Paragraph({
  children: [new TextRun({
    text: "A handful of small habits make a bigger difference than any single meal choice. None of this is complicated, but getting these locked in is what actually moves your numbers.",
    size: 22, font: FONT
  })],
  spacing: { after: 280, line: 320 }
}));

// Five tip paragraphs with bolded lead-ins
const tips = [
  ["Cut out liquid sugar entirely.", "Soda, sweet tea, juice, Gatorade, and flavored coffee creamer all hit your bloodstream faster than anything you could eat. Swapping them for water, black coffee, unsweetened tea, or sparkling water is the single biggest thing you can do for your A1C. On hot days when you're sweating through your shirt, grab a zero-sugar electrolyte packet (LMNT or Liquid IV Sugar Free) instead of Gatorade."],
  ["Flip every package before it goes in the cart.", "If the sugar content is over 3 grams per serving on any meat, sauce, or marinade, put it back. That one habit catches most of the traps. \"Honey BBQ\" wings, \"teriyaki\" anything, and \"maple\" sausage are all sugar wearing a meat costume."],
  ["Never eat fruit alone.", "Pair it with something fatty or with protein: apple with cheese, berries with Greek yogurt, or orange with almonds. The fat and protein slow sugar absorption and flatten the spike. Every meal in this plan already does this; just remember the rule when you're snacking on your own."],
  ["Real protein at every meal.", "Four ounces minimum, more if you're hungry. Protein keeps you full, keeps blood sugar steady through the afternoon, and protects muscle. You'll naturally hit 160 to 200 grams a day on this plan."],
  ["Grab a cheap glucose meter.", "They run about 20 dollars at Kroger. Test before you eat and two hours after, for just one week. You'll see exactly which foods spike you and which don't, and that kind of feedback sticks way better than general advice ever does."]
];

tips.forEach(([lead, body]) => {
  children.push(new Paragraph({
    children: [
      new TextRun({ text: lead + " ", bold: true, size: 22, font: FONT }),
      new TextRun({ text: body, size: 22, font: FONT })
    ],
    spacing: { after: 240, line: 320 }
  }));
});

// "How to Use This" explanation
children.push(h2("How to Use This"));
children.push(new Paragraph({
  children: [new TextRun({
    text: "The Shopping List is organized the way Kroger is laid out, so you can walk it top to bottom and be out in about half an hour. The Weekly Meal Plan lays out seven days with calories and sugar totals next to each meal so you always know what you're looking at. The Recipes give you step-by-step for everything on the plan. None take more than 20 minutes, and most are faster. The Full Kroger Reference is there for when you want to swap something out. Every cheese, protein, and prepared meal listed is diabetic-friendly, so you can mix and match freely.",
    size: 22, font: FONT
  })],
  spacing: { after: 320, line: 320 }
}));

// "A Few More Things That Help" wrap-up
children.push(h2("A Few More Things That Help"));

const morePoints = [
  "Cook once, eat twice. When you're making chicken thighs Wednesday, throw a couple extra in the air fryer so Friday's lunch is already handled. Sunday nights are your friend for prep: shred the rotisserie chicken, hard-boil half a dozen eggs, maybe pack a couple of lunch boxes, and the week basically runs itself.",
  "Salt and fat aren't the enemy. Olive oil, butter, full-fat cheese, and avocado are all good. The real villains are sugar and refined carbs (bread, rice, pasta, cereal, crackers, anything breaded). When you feel like snacking between meals, a hard-boiled egg, a string cheese, a handful of almonds, or some deli meat rolled around cheese will hold you over. Skip the chips.",
  "If you miss a meal or go off-plan one day, don't let it roll into two. Just pick it back up at the next meal. This plan works because of consistency, not perfection."
];

morePoints.forEach(text => {
  children.push(new Paragraph({
    children: [new TextRun({ text, size: 22, font: FONT })],
    spacing: { after: 220, line: 320 }
  }));
});

// Italic caveat - important for tone AND subtle liability protection
children.push(new Paragraph({
  children: [new TextRun({
    text: "One last thought: everyone's body works a little differently. If there are foods you react to or just don't tolerate well, swap them out freely using the reference list in Section 4. And since diabetes often comes with medications or other things going on that affect how you eat, it's always worth running any big changes past your doctor first. They know the full picture in a way a weekly meal plan never can. Think of this as a friendly starting point, not medical advice.",
    size: 22, font: FONT, italics: true
  })],
  spacing: { after: 320, line: 320 }
}));

children.push(new Paragraph({ children: [new PageBreak()] }));

// ---------------------------------------------------------------------------
// SECTION 1: WEEKLY KROGER SHOPPING LIST
// ---------------------------------------------------------------------------

children.push(h1("1. Weekly Kroger Shopping List"));
children.push(note("Aim for 30 to 45 min total. Stick to the perimeter (meat, produce, dairy) and a couple of pantry pickups. Estimated total: $90 to $120 for the week."));

// Customize shopping list by store section
// Each section: h2 header + note about weekly quantity + bullet items in "Item: qty" format

children.push(h2("MEAT"));
children.push(note("About 6 to 7 lbs of raw meat total for the week, enough for 7 dinners plus a couple of leftover lunches."));
children.push(bullet("Tyson Buffalo Wings (frozen): 2 bags (48 oz total, Mon and Fri at 10 to 12 wings each)"));
children.push(bullet("Johnsonville Brats or Italian Sausage: 1 pack (5 links, 3 for Tue dinner, 2 backup)"));
children.push(bullet("Boneless skinless chicken thighs: family pack (about 2.5 lbs, 6 thighs)"));
children.push(bullet("Sirloin steak bites or tips: 1.5 lbs"));
children.push(bullet("Ribeye or NY strip: 1 steak (10 to 12 oz)"));
children.push(bullet("Rotisserie chicken (deli): 1 whole (covers Sun dinner plus Sun and Wed lunches)"));
children.push(bullet("Bacon: 2 packs (12 oz each, about 3 strips across 4 breakfasts)"));

children.push(h2("DELI"));
children.push(note("About 1.25 lb combined meat for lunchbox charcuterie. Roughly 4 oz per lunch, 5 lunches."));
children.push(bullet("Sliced turkey (low sodium): 0.5 lb"));
children.push(bullet("Salami or prosciutto: 0.5 lb"));
children.push(bullet("Sliced ham or roast beef: 0.25 lb"));

children.push(h2("CHEESE"));
children.push(note("Pick 2 cheeses totaling 600 to 700g (20 to 24 oz). One hard (cheddar), one soft (brie, gouda). You're eating about 3 oz a day across lunch and breakfast."));
children.push(bullet("Sharp cheddar block: 12 oz"));
children.push(bullet("Brie, gouda, or manchego: 8 oz"));
children.push(bullet("String cheese or Babybel (snack): 1 pack"));

children.push(h2("DAIRY"));
children.push(note("24 eggs for the week. 1 tub Greek yogurt plus 1 tub cottage cheese covers the non-egg breakfasts."));
children.push(bullet("Eggs: 24-count (or 18 plus a spare 6-pack)"));
children.push(bullet("Butter (unsalted or salted): 1 stick or tub"));
children.push(bullet("Fage 5% Greek yogurt: 32 oz tub"));
children.push(bullet("Good Culture 4% cottage cheese: 1 tub (16 oz)"));

children.push(h2("PRODUCE"));
children.push(note("Avocado does heavy calorie lifting here, about half an avocado in most meals. Five total for the week. Greens and veg stay the same."));
children.push(bullet("Bagged spring mix or romaine: 2 bags"));
children.push(bullet("Cherry tomatoes: 1 pint"));
children.push(bullet("Cucumbers: 2"));
children.push(bullet("Bell peppers: 3-pack"));
children.push(bullet("Avocados: 5 (some ripe, some firmer to ripen through the week)"));
children.push(bullet("Strawberries or blueberries: 1 container"));
children.push(bullet("Apples (small): 2 to 3"));
children.push(bullet("Lemon: 1"));

children.push(h2("FROZEN"));
children.push(note("1 bag of steamable broccoli covers 2 to 3 side portions. Frozen berries as a backup if fresh spoils."));
children.push(bullet("Steamable broccoli: 1 bag"));
children.push(bullet("Frozen mixed berries (backup): 1 bag"));

children.push(h2("JARRED / PANTRY"));
children.push(note("Buy once, lasts 4 to 8 weeks. Skip this section if already stocked. Peanut butter and almonds get used across breakfasts and lunches, so don't skip those."));
children.push(bullet("Kalamata or mixed olives: 1 jar"));
children.push(bullet("Pepperoncini or banana peppers: 1 jar"));
children.push(bullet("Olive oil (if out): 1 bottle"));
children.push(bullet("Frank's RedHot or hot sauce: 1 bottle"));
children.push(bullet("Slap Ya Mama or Tony Chachere's seasoning: 1 jar"));
children.push(bullet("Almonds (unsalted, snack): 1 large bag (heavy weekly use)"));
children.push(bullet("Natural peanut butter (only ingredients: peanuts, salt): 1 jar"));

children.push(new Paragraph({ children: [new PageBreak()] }));

// ---------------------------------------------------------------------------
// SECTION 2: WEEKLY MEAL PLAN TABLE
// ---------------------------------------------------------------------------

children.push(h1("2. Weekly Meal List with Calories and Sugar"));
children.push(note(`Estimates are per-serving, approximate. Target about ${CONFIG.dailyCalorieTarget} calories a day on average. Sugar targets: under 10g breakfast, under 5g lunch, under 5g dinner. Bump portions further on heavy workdays; ease off on rest days.`));

// Each meal cell contains: description (line 1), then cal/sugar (line 2 in italic smaller)
// Customize these rows for different meal plans / calorie targets
const mealRows = [
  ["Day", "Breakfast", "Lunch", "Dinner", "Daily Totals"],
  ["Mon",
   "4 scrambled eggs + 3 bacon + half avocado\n~620 cal / 1g sugar",
   "4oz turkey + 3oz cheddar + olives + cucumber + tomato + olive oil + 1oz almonds\n~860 cal / 3g sugar",
   "10 to 12 Tyson Buffalo wings + bagged salad + olive oil + half avocado\n~1,050 cal / 2g sugar",
   "~2,530 cal / 6g sugar"],
  ["Tue",
   "1.5 cups Fage 5% + half cup strawberries + 1oz almonds + 1 tbsp PB\n~590 cal / 11g sugar",
   "4oz salami + 3oz brie + pepperoncini + spring mix + bell pepper + olive oil\n~860 cal / 3g sugar",
   "3 Johnsonville brats + bagged salad + olive oil + half avocado\n~1,060 cal / 4g sugar",
   "~2,510 cal / 18g sugar"],
  ["Wed",
   "4 eggs + 2 slices cheddar + half avocado + 2 bacon\n~700 cal / 1g sugar",
   "5oz rotisserie chicken + 2oz cheese cubes + cherry tomatoes + olives + 1oz nuts\n~800 cal / 2g sugar",
   "3 air-fryer chicken thighs + bagged salad + olive oil + half avocado\n~950 cal / 2g sugar",
   "~2,450 cal / 5g sugar"],
  ["Thu",
   "1.5 cups cottage cheese + half cup blueberries + 1oz almonds\n~520 cal / 11g sugar",
   "3oz ham + 2oz turkey + 3oz cheddar + cucumber + olives + small apple + 1oz almonds\n~900 cal / 16g sugar",
   "8 to 10oz steak bites + steamable broccoli + butter + half avocado\n~1,050 cal / 3g sugar",
   "~2,470 cal / 30g sugar"],
  ["Fri",
   "4 eggs + 3 bacon + half avocado\n~620 cal / 1g sugar",
   "2 leftover chicken thighs + 2oz brie + bell pepper + olives + olive oil\n~850 cal / 2g sugar",
   "10 to 12 Tyson Buffalo wings + bagged salad + olive oil\n~1,000 cal / 2g sugar",
   "~2,470 cal / 5g sugar"],
  ["Sat",
   "5 eggs + 3 bacon + 1 slice cheddar\n~680 cal / 1g sugar",
   "4oz prosciutto + 3oz cheddar + olives + tomato + cucumber + olive oil + 1oz nuts\n~920 cal / 3g sugar",
   "10oz ribeye + bagged salad + olive oil + half avocado\n~1,100 cal / 2g sugar",
   "~2,700 cal / 6g sugar"],
  ["Sun",
   "1.5 cups Fage 5% + half cup berries + 1oz almonds + 1 tbsp PB\n~620 cal / 11g sugar",
   "5oz rotisserie chicken + 3oz cheese + olives + spring mix + olive oil\n~850 cal / 2g sugar",
   "5oz leftover rotisserie chicken + steamable broccoli + butter + half avocado\n~820 cal / 3g sugar",
   "~2,290 cal / 16g sugar"],
];

const mealWidths = [700, 2200, 2200, 2200, 2060];

// Build the meal table cell-by-cell because each cell has multi-line content
// with different formatting on each line (bold for meal, italic small for cal/sugar)
const mealTable = new Table({
  width: { size: 9360, type: WidthType.DXA },
  columnWidths: mealWidths,
  rows: mealRows.map((row, rowIdx) => new TableRow({
    children: row.map((text, colIdx) => {
      const lines = text.split('\n');
      return new TableCell({
        borders,
        width: { size: mealWidths[colIdx], type: WidthType.DXA },
        // Only header row gets shading - no alternating row shading
        shading: rowIdx === 0 ? { fill: SWISS_COFFEE, type: ShadingType.CLEAR } : undefined,
        margins: { top: 80, bottom: 80, left: 100, right: 100 },
        children: lines.map((line, lineIdx) => new Paragraph({
          children: [new TextRun({
            text: line,
            bold: rowIdx === 0 || (lineIdx === 1 && rowIdx > 0),
            color: TEXT_COLOR,
            font: FONT,
            size: rowIdx === 0 ? 18 : (lineIdx === 1 ? 16 : 18),
            italics: lineIdx === 1 && rowIdx > 0
          })],
          spacing: { after: 20 }
        }))
      });
    })
  }))
});

children.push(mealTable);
children.push(blank());

children.push(h3("Weekly Nutrition Summary (approximate)"));
children.push(bullet(`Weekly average: about ${CONFIG.dailyCalorieTarget} calories a day, about 12g sugar a day`));
children.push(bullet("This plan sits just slightly below maintenance, which is a gentle way to support better blood sugar without losing weight."));
children.push(bullet("Protein range: 160 to 200 grams a day, plenty for holding onto muscle and staying full"));
children.push(bullet("Net carbs: under 35g most days, this is what moves A1C"));
children.push(bullet("Highest-calorie day: Saturday ribeye dinner. Lightest: Sunday, a built-in rest day."));
children.push(bullet("Feeling hungry? Add half an avocado, an extra egg at breakfast, or another handful of nuts. Don't add carbs."));

children.push(new Paragraph({ children: [new PageBreak()] }));

// ---------------------------------------------------------------------------
// SECTION 3: RECIPES
// ---------------------------------------------------------------------------
// Pull full content from references/recipes-library.md when building
// Below is an abbreviated version; expand with the library file contents

children.push(h1("3. Recipes (all under 20 minutes)"));

children.push(h2("Breakfasts"));

children.push(h3("Scrambled Eggs + Air Fryer Bacon"));
children.push(note("Prep 2 min | Cook 8 min | Total 10 min | ~620 cal"));
children.push(bullet("Lay 3 strips bacon in air fryer basket. 400 degrees for 8 min."));
children.push(bullet("While bacon cooks: crack 4 eggs in a bowl, whisk with pinch of salt."));
children.push(bullet("Melt pat of butter in skillet over medium. Pour in eggs, stir gently until just set, about 3 min."));
children.push(bullet("Plate eggs, top with bacon. Add half a sliced avocado (important: adds calories and healthy fat)."));

// ... (add other breakfast, lunch, dinner recipes by pulling from recipes-library.md) ...

children.push(new Paragraph({ children: [new PageBreak()] }));

// ---------------------------------------------------------------------------
// SECTION 4: FULL KROGER REFERENCE
// ---------------------------------------------------------------------------
// Pull full content from references/kroger-products.md when building

children.push(h1("4. Exhaustive Kroger Reference List"));
children.push(note("Rotate through these when you get bored of the core plan. Everything listed is low-sugar and diabetic-safe when used in reasonable portions. Each section has its own weekly quantity guidance at the top."));

// Add all six category sections (Cheese, Meat and Protein, Fruit, Prepared Meals, Pantry, Drinks)
// from references/kroger-products.md

// ============================================================================
// BUILD THE DOCUMENT
// ============================================================================

const doc = new Document({
  styles: {
    default: { document: { run: { font: FONT, size: 22, color: TEXT_COLOR } } },
    paragraphStyles: [
      { id: "Heading1", name: "Heading 1", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: { size: 36, bold: true, font: FONT, color: TEXT_COLOR },
        paragraph: { spacing: { before: 280, after: 180 }, outlineLevel: 0 } },
      { id: "Heading2", name: "Heading 2", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: { size: 28, bold: true, font: FONT, color: TEXT_COLOR },
        paragraph: { spacing: { before: 220, after: 120 }, outlineLevel: 1 } },
      { id: "Heading3", name: "Heading 3", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: { size: 24, bold: true, font: FONT, color: TEXT_COLOR },
        paragraph: { spacing: { before: 160, after: 80 }, outlineLevel: 2 } },
    ]
  },
  numbering: {
    config: [
      { reference: "bullets",
        levels: [{
          level: 0, format: LevelFormat.BULLET, text: "•", alignment: AlignmentType.LEFT,
          style: { paragraph: { indent: { left: 720, hanging: 360 } } }
        }]
      }
    ]
  },
  sections: [{
    properties: {
      page: {
        size: { width: 12240, height: 15840 },         // US Letter
        margin: { top: 1080, right: 1080, bottom: 1080, left: 1080 }  // 0.75 inch
      }
    },
    children: children
  }]
});

Packer.toBuffer(doc).then(buffer => {
  // Ensure output directory exists
  const outputDir = CONFIG.outputPath.substring(0, CONFIG.outputPath.lastIndexOf('/'));
  if (outputDir && !fs.existsSync(outputDir)) {
    fs.mkdirSync(outputDir, { recursive: true });
  }
  fs.writeFileSync(CONFIG.outputPath, buffer);
  console.log(`Built successfully: ${CONFIG.outputPath}`);
});
