# OneAccord Training Portal — Shell (v0)

This is the first working piece of the training portal: a nav + a block-rendering
engine that can step a learner through a training one piece at a time, with the
four block types the Zoho CRM Coach training needs (`text`, `video`, `activity`,
`question`). It's static HTML/JS/CSS — no build step, no framework, deploys as-is.

**What's real right now:** Lessons 1 and 2 of the Zoho CRM Training Coach are
fully built out from the master guide content and playable end to end. Lessons
3–9, and all of Sales Training, show in the nav as "pending" — content for those
comes from Jesse's per-training chats, per the project plan.

## Try it locally

No install needed — any static file server works:

```
cd training-portal
python3 -m http.server 8080
```

Then open `http://localhost:8080`.

## Deploy (GitHub + Netlify, per the plan)

1. Push this folder to a new GitHub repo (e.g. `oneaccord/training-portal`).
2. In Netlify, "Add new site" → "Import an existing project" → pick that repo.
   Build command: none. Publish directory: `.` (already set in `netlify.toml`).
3. Netlify gives you a URL (e.g. `oneaccord-training.netlify.app`, or attach a
   custom domain). That URL is what goes into SharePoint's Embed web part.

## Embedding in SharePoint

On the training page in SharePoint: **Add a web part → Embed**, paste the
Netlify URL, save. If it doesn't render, an admin needs to add the Netlify
domain to SharePoint's embed allow-list (Jesse has the rights to do this —
see "Allow or restrict the ability to embed content on SharePoint pages" in
SharePoint admin settings).

## Adding a new training or lesson (today, by hand)

Until the SharePoint lists are wired up (see below), adding content means:

1. Add an entry to `data/trainings-index.json` (a lesson, or a whole new
   training) with `"status": "ready"`.
2. Add a matching JSON file under `data/trainings/<training-id>/<lesson-id>.json`
   following the shape in `lesson-1.json` / `lesson-2.json`.
3. Push to GitHub — Netlify redeploys automatically.

This is exactly what the SharePoint "Trainings" list will do once it's wired
up: the index list becomes `trainings-index.json`'s replacement, so adding a
row there does what adding a JSON entry does here, without touching code.

## The two SharePoint lists (Foundation phase — still to create)

These don't exist yet. Create them as standard SharePoint **Lists** (not
document libraries) on the training site. Column names below are suggestions;
what matters is that each concept has a column.

### List 1: `Trainings` (the index — drives the nav)

| Column | Type | Notes |
| --- | --- | --- |
| Title | Single line text | Lesson/module title shown in the nav |
| TrainingId | Single line text | Groups lessons under one training, e.g. `zoho-crm-coach` |
| LessonId | Single line text | Matches the JSON filename, e.g. `lesson-3` |
| Category | Single line text | The training's display name, e.g. "Zoho CRM Training Coach" |
| Order | Number | Nav display order |
| Status | Choice: Ready / Pending / Retired | Whether it shows as clickable |
| ContentUrl | Single line text (URL) | Where the lesson's content package (JSON) lives |

### List 2: `Completions` (tracking — drives "who's finished" and percent-complete)

| Column | Type | Notes |
| --- | --- | --- |
| Person | Person or Group | Who completed it |
| TrainingId | Single line text | |
| LessonId | Single line text | |
| CompletedOn | Date and time | |
| Score | Number | Optional — only used by AI-graded assessments (Sales Training) |

Once these exist, Phase 3's Netlify Function reads `Trainings` (via Microsoft
Graph) to build the index instead of `trainings-index.json`, and writes a row
to `Completions` instead of `localStorage`. Nothing in `app.js`'s rendering
logic needs to change — only the two functions marked `TODO(phase 3)` at the
bottom of `js/app.js`.
