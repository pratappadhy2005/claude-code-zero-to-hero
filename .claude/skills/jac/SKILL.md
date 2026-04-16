---
name: student-test-analyser
description: >
  Analyses a student's test result sheet (uploaded image or typed data) and produces:
  (1) a detailed visual diagnostic report showing performance by topic, right/wrong per question,
  and a priority action plan, and (2) a downloadable .docx practice worksheet targeting the
  exact weak topics identified. Use this skill whenever a parent or student shares test results,
  score sheets, a percentage score with topic breakdowns, or asks how to improve in a specific
  subject or test (e.g. selective school entry, OC, NAPLAN, general ability, EduTest, ACER).
  Also trigger when someone asks to "set milestones", "make a study plan", or "create practice
  questions" based on a test or exam result — even if they haven't uploaded a file yet.
---

# Student test analyser & worksheet generator

## Overview

This skill does two things in sequence:

1. **Diagnostic analysis** — reads the student's test result (image or typed data), identifies
   wrong answers, groups them by topic, and renders an interactive visual report with a
   priority action plan.

2. **Practice worksheet** — generates a downloadable `.docx` file with ~30 targeted questions
   covering the weak topics found in step 1, complete with an answer key and scoring guide.

---

## Step 0 — Gather inputs

Before doing anything, make sure you have:

| Input | How to get it |
|---|---|
| Student's score / result sheet | Uploaded image, typed list, or percentage + topic list |
| Year level / grade | Ask if not mentioned |
| School or test name | Ask if not mentioned — affects question style |
| Subject area | General ability, English, Maths, Science, etc. |

If the user uploads an image of a result sheet, read it carefully:
- `Q.No` — question number
- `Resp` — what the student answered (O = correct, X = wrong)
- `Ans` — correct answer
- `Percent` — how many students in cohort answered correctly (difficulty indicator)
- `Topic` — the skill area tested

If no image is uploaded but a score and topic list are given, work from that.

---

## Step 1 — Build the diagnostic visual

Use the `visualize:show_widget` tool to render an interactive HTML dashboard. Include:

### Metric summary cards (top row)
- Total score (e.g. 17/30)
- Number of wrong answers
- Weakest topic (0% or lowest)
- Strongest topic (100% or highest)

### Topic performance table
For each topic: name, correct/total, percentage bar coloured by urgency, priority tag.

Colour coding:
- Red / "Priority" → below 50%
- Amber / "Improve" → 50–79%
- Green / "Good" → 80%+

Sort topics from lowest to highest percentage.

### Question-by-question grid
Show all questions as small cells. Green dot = correct, red dot = wrong. Include topic label.

### Priority action plan
4 action cards targeting the top weak topic clusters. Be specific:
- Name the exact question numbers that were wrong
- State a concrete daily practice action (e.g. "10 analogy pairs each morning")
- Explain *why* the topic is weak, not just *that* it is

---

## Step 2 — Generate the practice worksheet (.docx)

Read `/mnt/skills/public/docx/SKILL.md` before writing any code.

Install dependency if needed:
```bash
npm list -g docx || npm install -g docx
```

### Worksheet structure

```
Cover page
  Student name / date / score fields
  Instructions box

Section per weak topic  (prioritised — worst topics first)
  Section header with colour-coded priority badge
  6–8 questions per section
  Multiple-choice (A/B/C/D) laid out in a 4-column table

Answer key page  (separate, clearly marked for parent/teacher)
  Q number | Correct answer | Plain-English explanation

Scoring guide
  What to do at each score band
```

### Question quality rules
- Questions must match the year level (Grade 5 = approx 10–11 year old vocabulary)
- Use Australian English spelling (colour, programme, recognise)
- Multiple choice options must be plausible — avoid obviously wrong distractors
- For analogy questions: always use the format "X is to Y as ___ is to ___"
- For science vocabulary: test word meanings, not recall of facts
- For palindromes: always include a worked example in the section subheading
- For idioms/colloquialisms: use Australian and British expressions where possible

### docx visual conventions
- A4 page size (11906 × 16838 DXA), 1260 DXA margins
- Section headers: coloured table row (red/amber/green fill matching priority)
- MCQ options: 4-column borderless table, evenly spaced
- Answer key header row: dark blue fill (#1A5276) with white text
- Alternating row shading on answer key for readability
- Cover page: name/date/score/time fields in a 4-column table
- Instruction box: light blue shading (#EAF4FF)

### Code pattern (Node.js / docx-js)

```javascript
const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  AlignmentType, BorderStyle, WidthType, ShadingType, LevelFormat, PageBreak
} = require('docx');
const fs = require('fs');

// Build sections array from weak topics
// Each section: sectionHeader() + subtext() + questions + divider()
// Final page: answer key table with explanation column

Packer.toBuffer(doc).then(buf => {
  fs.writeFileSync('/home/claude/practice_worksheet.docx', buf);
});
```

Save to `/home/claude/practice_worksheet.docx`, then copy to `/mnt/user-data/outputs/` and
call `present_files` to share with the user.

---

## Step 3 — Present results

After both the visual and the file are ready:

1. Show the visual diagnostic first (already rendered inline)
2. Share the `.docx` file via `present_files`
3. Give a short plain-English summary:
   - Which 2–3 topics to fix first and why
   - One concrete thing to do today
   - Encourage: frame 56% → 85% as achievable in ~16 weeks with the plan

---

## Milestone framework (use when parent asks about improvement plan)

If the student's current score is known, map to this 3-milestone structure:

| Milestone | Target | Focus | Weeks |
|---|---|---|---|
| 1 | Current + 9% | Understand each question type; untimed practice | 1–4 |
| 2 | Current + 19% | Speed + accuracy; timed sessions; error journal | 5–10 |
| 3 | 85%+ | Exam simulation; full timed tests; consistency | 11–16 |

Weekly routine template (45 min/day):
- Mon: weakest verbal topic — 20 questions
- Tue: weakest non-verbal / pattern topic — timed 15 Qs in 20 min
- Wed: maths / quantitative drills
- Thu: reading comprehension passage (NAPLAN-style)
- Fri: mixed mini-test across all topics, mark and analyse
- Weekend (every 2nd week): full practice test

---

## Subject-specific notes

### General ability (EduTest / ACER style — Victoria selective schools)
Common topics: Analogies, Synonyms, Antonyms, Definitions, Science vocab, Humanities,
Mathematics, Colloquialisms, Idioms, Suffixes, Prefixes, Palindromes, Superlatives,
Geography, Spelling

Key insight: the `Percent` column shows cohort difficulty. If a student gets a question
wrong that 80%+ of students got right, it is a priority fix — not a hard question.

### NAPLAN
Focus on: reading comprehension, numeracy, language conventions (spelling, grammar, punctuation)

### OC / selective entry (NSW)
Focus on: mathematical reasoning, reading comprehension, thinking skills

---

## Error handling

- If the uploaded image is blurry or partially cut off: ask the user to confirm the
  wrong answers before proceeding
- If fewer than 5 wrong answers are found: still generate the worksheet but note that
  the student is performing well and shift focus to speed and timed practice
- If no image is provided but a topic list is given: generate the worksheet from the
  topics listed, and note assumptions made