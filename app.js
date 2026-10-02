/**
 * OneAccord Training Portal — shell engine
 * ------------------------------------------------
 * Renders a nav from a "trainings index" and steps a learner through a
 * training's content blocks one at a time, gating progress on hands-on
 * confirmations and answered questions.
 *
 * DATA SOURCE: the SharePoint "Trainings" list, read via the
 * /.netlify/functions/get-trainings function (Microsoft Graph, app-only
 * auth). Falls back to the static /data/trainings-index.json file if that
 * call fails, so the portal still works if SharePoint/Graph is unreachable.
 *
 * LEARNER IDENTITY: a lightweight "email gate" (see showEmailGate()) asks
 * the learner for their work email once, stores it in localStorage, and
 * sends it with every completion read/write so progress can be tied to a
 * person in the SharePoint "Training Completions" list without a full
 * sign-in flow.
 *
 * COMPLETION TRACKING: writes go to /.netlify/functions/record-completion
 * (one row per training+lesson+email in the SharePoint "Training
 * Completions" list); reads come from /.netlify/functions/get-completions,
 * filtered by the learner's email, and are cached in `state.completions`
 * for the session.
 *
 * BLOCK TYPES SUPPORTED (the four the first pilot needs — see README):
 *   text      — static explanatory copy
 *   video     — an embedded video (SharePoint stream link or a file)
 *   activity  — a hands-on "find it now" step the learner confirms in person
 *   question  — a fixed-answer multiple-choice or true/false check
 *
 * ADDED FOR SALES TRAINING (see "AI-coached blocks" below):
 *   links       — copy plus a row of link buttons (resources, SharePoint)
 *   scenario    — free-text answer, AI-graded against a rubric
 *   deliverable — a Game Plan section, AI-reviewed (approved / needs revision)
 *   reflection  — optional free-text, short AI coaching reply
 *   gameplan    — compiles saved deliverables into a downloadable plan
 *   video       — now also supports `note`, `links` and `confirmText`
 */

const state = {
  index: null,
  currentTrainingId: null,
  currentLessonId: null,
  lessonBlocks: [],
  blockCursor: 0,
  expandedTrainings: new Set(), // training ids whose lesson list is open in the nav
  completions: new Set(), // `${trainingId}:${lessonId}` keys the current learner has completed
  personEmail: null,
};

const navEl = document.getElementById("training-list");
const mainEl = document.getElementById("portal-main");

init();

async function init() {
  try {
    let index = null;
    try {
      const res = await fetch("/.netlify/functions/get-trainings");
      if (!res.ok) throw new Error(`get-trainings returned ${res.status}`);
      index = await res.json();
    } catch (err) {
      console.warn("Falling back to static trainings index", err);
      const res = await fetch("data/trainings-index.json");
      index = await res.json();
    }
    state.index = index;
    // Default to the first training open so the nav isn't empty on first load.
    if (state.index.trainings.length) {
      state.expandedTrainings.add(state.index.trainings[0].id);
    }

    state.personEmail = getStoredEmail();
    if (state.personEmail) {
      await loadCompletions();
    }

    renderNav();

    if (!state.personEmail) {
      showEmailGate();
    }
  } catch (err) {
    navEl.innerHTML = `<div style="padding:0 20px;color:#f5b7b1;">Couldn't load the training list.</div>`;
    console.error(err);
  }
}

/* ---------------- Learner identity (email gate) ----------------
 * We ask for a work email once (no password — this isn't meant to be a
 * security boundary, just a way to attribute completions to a person in
 * the SharePoint list). The email is stored in localStorage so returning
 * learners aren't asked again on this device/browser.
 */

const EMAIL_STORAGE_KEY = "oa-training-email";

function getStoredEmail() {
  try {
    return localStorage.getItem(EMAIL_STORAGE_KEY) || null;
  } catch (err) {
    return null;
  }
}

function storeEmail(email) {
  try {
    localStorage.setItem(EMAIL_STORAGE_KEY, email);
  } catch (err) {
    console.warn("Could not store email locally", err);
  }
}

function showEmailGate() {
  mainEl.innerHTML = `<div class="email-gate">
      <h2>Before you start</h2>
      <p>Enter your work email so we can track your training progress.</p>
      <div class="error" style="display:none;"></div>
      <input type="email" placeholder="you@oneaccord.co" autocomplete="email" />
      <button type="button" class="primary-btn">Continue</button>
    </div>`;

  const wrap = mainEl.querySelector(".email-gate");
  const input = wrap.querySelector("input");
  const errorEl = wrap.querySelector(".error");
  const submit = wrap.querySelector("button");

  const submitEmail = async () => {
    const email = input.value.trim().toLowerCase();
    if (!email || !email.includes("@") || !email.includes(".")) {
      errorEl.textContent = "Please enter a valid email address.";
      errorEl.style.display = "block";
      return;
    }
    submit.disabled = true;
    submit.textContent = "Loading…";
    state.personEmail = email;
    storeEmail(email);
    await loadCompletions();
    renderNav();
    mainEl.innerHTML = `<div class="portal-welcome"><h1>Welcome!</h1><p>Pick a lesson from the menu on the left to get started.</p></div>`;
  };

  submit.addEventListener("click", submitEmail);
  input.addEventListener("keydown", (e) => {
    if (e.key === "Enter") submitEmail();
  });
  input.focus();
}

async function loadCompletions() {
  if (!state.personEmail) return;
  try {
    const res = await fetch(`/.netlify/functions/get-completions?personEmail=${encodeURIComponent(state.personEmail)}`);
    if (!res.ok) throw new Error(`get-completions returned ${res.status}`);
    const data = await res.json();
    state.completions = new Set(
      (data.completions || []).map((c) => `${c.trainingId}:${c.lessonId}`)
    );
  } catch (err) {
    console.warn("Could not load completions", err);
  }
}

/* ---------------- Nav ----------------
 * Each training is a collapsible section: a clickable header (with a
 * chevron) toggles its lesson list open/closed. `state.expandedTrainings`
 * tracks which sections are open across re-renders; loading a lesson always
 * opens that lesson's training so the active item is visible.
 */

function renderNav() {
  navEl.innerHTML = "";
  state.index.trainings.forEach((training) => {
    const isOpen = state.expandedTrainings.has(training.id);

    const heading = document.createElement("button");
    heading.type = "button";
    heading.className = "nav-training" + (isOpen ? " is-open" : "");
    heading.innerHTML = `<span class="nav-training__chevron">▸</span><span>${training.title}</span>`;
    heading.addEventListener("click", () => {
      if (state.expandedTrainings.has(training.id)) {
        state.expandedTrainings.delete(training.id);
      } else {
        state.expandedTrainings.add(training.id);
      }
      renderNav();
    });
    navEl.appendChild(heading);

    const list = document.createElement("div");
    list.className = "nav-lessons" + (isOpen ? "" : " is-collapsed");

    training.lessons.forEach((lesson) => {
      const item = document.createElement("div");
      const isPending = lesson.status === "pending";
      const isDone = isLessonComplete(training.id, lesson.id);
      const isActive = training.id === state.currentTrainingId && lesson.id === state.currentLessonId;
      item.className = "nav-lesson" + (isPending ? " is-pending" : "") + (isActive ? " is-active" : "");
      item.innerHTML = `<span class="nav-lesson__check">${isDone ? "✅" : "⬜"}</span><span>${lesson.title}</span>`;
      if (!isPending) {
        item.addEventListener("click", () => {
          if (!state.personEmail) {
            showEmailGate();
            return;
          }
          loadLesson(training, lesson);
        });
      } else {
        item.title = "Content coming soon";
      }
      list.appendChild(item);
    });

    navEl.appendChild(list);
  });
}

/* ---------------- Lesson loading ---------------- */

async function loadLesson(training, lesson) {
  state.currentTrainingId = training.id;
  state.currentLessonId = lesson.id;
  state.blockCursor = 0;
  state.expandedTrainings.add(training.id);

  mainEl.innerHTML = `<div class="lesson-header">
      <h1>${lesson.title}</h1>
      <div class="lesson-progress">${training.title}</div>
    </div>
    <div id="block-stream"></div>`;

  const contentUrl = lesson.contentUrl || `data/trainings/${training.id}/${lesson.id}.json`;
  const res = await fetch(contentUrl);
  const lessonDef = await res.json();
  state.lessonBlocks = lessonDef.blocks;
  state.lessonMeta = lessonDef;

  renderNav();
  renderNextBlock();
}

function getStream() {
  return document.getElementById("block-stream");
}

function renderNextBlock() {
  const stream = getStream();
  const idx = state.blockCursor;

  if (idx >= state.lessonBlocks.length) {
    renderLessonComplete();
    return;
  }

  const block = state.lessonBlocks[idx];
  const el = document.createElement("div");
  el.className = "block";

  if (block.type === "text") {
    el.classList.add("block-text");
    el.innerHTML = block.html;
    stream.appendChild(el);
    appendContinueButton(stream);
  } else if (block.type === "video") {
    el.classList.add("block-video");
    el.innerHTML = `<iframe src="${block.src}" allowfullscreen title="${block.title || "Training video"}"></iframe>`;
    stream.appendChild(el);
    // Optional extras (Sales Training): a note under the video and fallback links
    // (open in SharePoint / read the transcript) for when the embed won't play.
    if (block.note || (block.links && block.links.length)) {
      const extras = document.createElement("div");
      extras.className = "block video-extras";
      extras.innerHTML =
        (block.note ? `<div class="callout">${block.note}</div>` : "") +
        renderLinks(block.links);
      stream.appendChild(extras);
    }
    if (block.confirmText) {
      appendConfirmButton(stream, block.confirmText);
    } else {
      appendContinueButton(stream);
    }
  } else if (block.type === "links") {
    el.classList.add("block-text");
    el.innerHTML = (block.html || "") + renderLinks(block.links);
    stream.appendChild(el);
    appendContinueButton(stream);
  } else if (block.type === "scenario") {
    renderScenarioBlock(el, block);
    stream.appendChild(el);
  } else if (block.type === "deliverable") {
    renderDeliverableBlock(el, block);
    stream.appendChild(el);
  } else if (block.type === "reflection") {
    renderReflectionBlock(el, block);
    stream.appendChild(el);
  } else if (block.type === "gameplan") {
    renderGamePlanBlock(el, block);
    stream.appendChild(el);
    appendContinueButton(stream);
  } else if (block.type === "activity") {
    el.classList.add("block-activity");
    el.innerHTML = `<div class="block-activity__label">🔍 ${block.label || "Find it now"}</div>
      <div>${block.instructions}</div>`;
    stream.appendChild(el);
    appendConfirmButton(stream, block.confirmText || "I've completed this — continue");
  } else if (block.type === "question") {
    renderQuestionBlock(el, block);
    stream.appendChild(el);
    // continue button is appended once the question is answered (see renderQuestionBlock)
  }

  stream.scrollTop = stream.scrollHeight;
}

function appendContinueButton(stream) {
  const btn = document.createElement("button");
  btn.className = "primary-btn";
  btn.textContent = "Continue";
  btn.addEventListener("click", () => {
    btn.remove();
    state.blockCursor += 1;
    renderNextBlock();
  });
  stream.appendChild(btn);
}

function appendConfirmButton(stream, label) {
  const btn = document.createElement("button");
  btn.className = "primary-btn";
  btn.textContent = label;
  btn.addEventListener("click", () => {
    btn.remove();
    state.blockCursor += 1;
    renderNextBlock();
  });
  stream.appendChild(btn);
}

function renderQuestionBlock(el, block) {
  el.classList.add("block-question");
  const prompt = document.createElement("div");
  prompt.className = "block-question__prompt";
  prompt.textContent = block.prompt;
  el.appendChild(prompt);

  const options = block.format === "true_false" ? ["True", "False"] : block.options;

  options.forEach((optionText, i) => {
    const btn = document.createElement("button");
    btn.className = "option-btn";
    btn.textContent = optionText;
    btn.addEventListener("click", () => {
      const allButtons = el.querySelectorAll(".option-btn");
      allButtons.forEach((b) => (b.disabled = true));

      const isCorrect = i === block.correctIndex;
      btn.classList.add(isCorrect ? "is-correct" : "is-incorrect");
      if (!isCorrect) {
        allButtons[block.correctIndex].classList.add("is-correct");
      }

      const feedback = document.createElement("div");
      feedback.className = "feedback " + (isCorrect ? "is-correct" : "is-incorrect");
      feedback.textContent = isCorrect
        ? block.feedbackCorrect || "Correct."
        : block.feedbackIncorrect || "Not quite — see the correct answer above.";
      el.appendChild(feedback);

      appendContinueButton(getStream());
    });
    el.appendChild(btn);
  });
}

/* ---------------- AI-coached blocks (Sales Training) ----------------
 * scenario    — free-text answer graded against a rubric by the `coach`
 *               Netlify Function; must reach passThreshold to continue.
 * deliverable — one section of the learner's Sales Game Plan; the coach
 *               returns Strengths / Recommendations / Approved or Needs Revision.
 * reflection  — optional; the coach replies with a short coaching note.
 * gameplan    — compiles every saved deliverable into one downloadable plan.
 *
 * The browser only sends ids + the learner's text. The function looks up the
 * rubric/criteria itself from the published lesson JSON, so it can't be used
 * as a general-purpose AI proxy. If the function is unreachable (e.g. the API
 * key isn't configured yet), each block falls back to a self-check so the
 * training still works.
 */

const COACH_URL = "/.netlify/functions/coach";

function escapeHtml(str) {
  return String(str == null ? "" : str)
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#39;");
}

function renderLinks(links) {
  if (!links || !links.length) return "";
  return `<div class="link-row">${links
    .map(
      (l) =>
        `<a class="link-btn" href="${l.url}" target="_blank" rel="noopener">${escapeHtml(l.label)}</a>`
    )
    .join("")}</div>`;
}

async function callCoach(payload) {
  const controller = new AbortController();
  const timer = setTimeout(() => controller.abort(), 60000);
  try {
    const res = await fetch(COACH_URL, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({
        trainingId: state.currentTrainingId,
        lessonId: state.currentLessonId,
        ...payload,
      }),
      signal: controller.signal,
    });
    if (!res.ok) throw new Error(`Coach returned ${res.status}`);
    return await res.json();
  } finally {
    clearTimeout(timer);
  }
}

function scoreKey(trainingId, lessonId, blockId) {
  return `oa-training-score:${trainingId}:${lessonId}:${blockId}`;
}

function gamePlanKey(trainingId, lessonId, blockId) {
  return `oa-gameplan:${trainingId}:${lessonId}:${blockId}`;
}

function saveLocal(key, value) {
  try {
    localStorage.setItem(key, JSON.stringify(value));
  } catch (err) {
    console.warn("Could not save locally", err);
  }
}

function loadLocal(key) {
  try {
    const raw = localStorage.getItem(key);
    return raw ? JSON.parse(raw) : null;
  } catch (err) {
    return null;
  }
}

function aiShell(el, label, promptHtml) {
  el.classList.add("block-ai");
  el.innerHTML = `<div class="block-ai__label">${label}</div>
    <div class="block-ai__prompt">${promptHtml}</div>`;
}

function renderScenarioBlock(el, block) {
  aiShell(el, "📝 Scenario", block.promptHtml);
  const threshold = block.passThreshold || Math.ceil(block.rubric.length * 0.75);

  const input = document.createElement("textarea");
  input.className = "ai-input";
  input.rows = 8;
  input.maxLength = 4000;
  input.placeholder = "Write your response here…";
  el.appendChild(input);

  const submit = document.createElement("button");
  submit.className = "primary-btn";
  submit.textContent = "Submit for coaching";
  el.appendChild(submit);

  const result = document.createElement("div");
  result.className = "ai-result";
  el.appendChild(result);

  let attempts = 0;

  submit.addEventListener("click", async () => {
    const answer = input.value.trim();
    if (answer.length < 40) {
      result.innerHTML = `<div class="feedback is-incorrect">Give it a bit more detail. Write it the way you'd actually say it.</div>`;
      return;
    }
    attempts += 1;
    submit.disabled = true;
    input.disabled = true;
    result.innerHTML = `<div class="ai-thinking">Your coach is reviewing your answer…</div>`;

    try {
      const res = await callCoach({ mode: "scenario", blockId: block.id, answer });
      const points = block.rubric
        .map((r, i) => {
          const p = (res.points && res.points[i]) || {};
          return `<li class="${p.met ? "is-met" : "is-missed"}"><span>${p.met ? "✅" : "⬜"}</span>
            <div><strong>${escapeHtml(r)}</strong>${p.note ? `<div class="rubric-note">${escapeHtml(p.note)}</div>` : ""}</div></li>`;
        })
        .join("");
      const passed = res.score >= threshold;
      result.innerHTML = `<div class="feedback ${passed ? "is-correct" : "is-incorrect"}">
          <strong>${res.score} of ${block.rubric.length} points${passed ? " — passed" : ` — ${threshold} needed to pass`}</strong></div>
        <ul class="rubric-list">${points}</ul>
        ${res.coaching ? `<div class="coach-note">${escapeHtml(res.coaching)}</div>` : ""}`;

      if (passed) {
        saveLocal(scoreKey(state.currentTrainingId, state.currentLessonId, block.id), {
          score: res.score,
          outOf: block.rubric.length,
          at: new Date().toISOString(),
        });
        appendContinueButton(getStream());
      } else {
        submit.disabled = false;
        input.disabled = false;
        submit.textContent = "Revise and resubmit";
      }
    } catch (err) {
      console.warn("Coach unavailable, falling back to self-check", err);
      result.innerHTML = `<div class="callout">AI coaching isn't available right now. Check your answer against these points, then continue:</div>
        <ul class="rubric-list">${block.rubric.map((r) => `<li><span>☐</span><div>${escapeHtml(r)}</div></li>`).join("")}</ul>`;
      appendConfirmButton(getStream(), "I've checked my answer against these points — continue");
    }
  });
}

function renderDeliverableBlock(el, block) {
  aiShell(el, `🗂️ ${escapeHtml(block.label || "Game Plan")}: ${escapeHtml(block.title)}`, block.promptHtml);
  const saved = loadLocal(gamePlanKey(state.currentTrainingId, state.currentLessonId, block.id));

  const inputs = {};
  block.fields.forEach((f) => {
    const wrap = document.createElement("label");
    wrap.className = "ai-field";
    wrap.innerHTML = `<span>${escapeHtml(f.label)}</span>`;
    const ta = document.createElement("textarea");
    ta.className = "ai-input";
    ta.rows = f.rows || 5;
    ta.maxLength = 3000;
    ta.placeholder = f.placeholder || "";
    if (saved && saved.fields && saved.fields[f.id]) ta.value = saved.fields[f.id];
    wrap.appendChild(ta);
    el.appendChild(wrap);
    inputs[f.id] = ta;
  });

  const submit = document.createElement("button");
  submit.className = "primary-btn";
  submit.textContent = "Submit for review";
  el.appendChild(submit);

  const result = document.createElement("div");
  result.className = "ai-result";
  el.appendChild(result);

  let attempts = 0;
  let continued = false;

  const save = (status) => {
    const fields = {};
    Object.keys(inputs).forEach((k) => (fields[k] = inputs[k].value.trim()));
    saveLocal(gamePlanKey(state.currentTrainingId, state.currentLessonId, block.id), {
      title: block.title,
      fields,
      labels: Object.fromEntries(block.fields.map((f) => [f.id, f.label])),
      status,
      at: new Date().toISOString(),
    });
    return fields;
  };

  const proceed = () => {
    if (continued) return;
    continued = true;
    appendContinueButton(getStream());
  };

  submit.addEventListener("click", async () => {
    const fields = save("draft");
    const empty = block.fields.filter((f) => (fields[f.id] || "").length < 15);
    if (empty.length) {
      result.innerHTML = `<div class="feedback is-incorrect">Please complete: ${empty.map((f) => escapeHtml(f.label)).join(", ")}.</div>`;
      return;
    }
    attempts += 1;
    submit.disabled = true;
    result.innerHTML = `<div class="ai-thinking">Your coach is reviewing your plan…</div>`;

    try {
      const res = await callCoach({ mode: "deliverable", blockId: block.id, fields });
      const approved = res.status === "approved";
      save(approved ? "approved" : "needs_revision");
      result.innerHTML = `<div class="feedback ${approved ? "is-correct" : "is-incorrect"}"><strong>${approved ? "Approved" : "Needs revision"}</strong></div>
        ${res.strengths && res.strengths.length ? `<p><strong>Strengths</strong></p><ul>${res.strengths.map((s) => `<li>${escapeHtml(s)}</li>`).join("")}</ul>` : ""}
        ${res.recommendations && res.recommendations.length ? `<p><strong>Recommendations</strong></p><ul>${res.recommendations.map((s) => `<li>${escapeHtml(s)}</li>`).join("")}</ul>` : ""}`;

      if (approved) {
        proceed();
      } else {
        submit.disabled = false;
        submit.textContent = "Revise and resubmit";
        // Don't trap anyone behind an over-strict review: after two tries they
        // can save their draft and move on (it's still in their Game Plan).
        if (attempts >= 2 && !el.querySelector(".secondary-btn")) {
          const skip = document.createElement("button");
          skip.className = "secondary-btn";
          skip.textContent = "Save my draft and continue";
          skip.addEventListener("click", () => {
            skip.remove();
            save("draft");
            proceed();
          });
          el.appendChild(skip);
        }
      }
    } catch (err) {
      console.warn("Coach unavailable, saving without review", err);
      save("saved");
      result.innerHTML = `<div class="callout">AI review isn't available right now, so your work was saved to your Game Plan without review. Check it against these points:</div>
        <ul>${(block.criteria || []).map((c) => `<li>${escapeHtml(c)}</li>`).join("")}</ul>`;
      proceed();
    }
  });
}

function renderReflectionBlock(el, block) {
  aiShell(el, "💭 Reflection (optional)", `<p>${escapeHtml(block.prompt)}</p>`);
  const input = document.createElement("textarea");
  input.className = "ai-input";
  input.rows = 4;
  input.maxLength = 2000;
  el.appendChild(input);

  const actions = document.createElement("div");
  actions.className = "lesson-nav-actions";
  const share = document.createElement("button");
  share.className = "primary-btn";
  share.textContent = "Share my reflection";
  const skip = document.createElement("button");
  skip.className = "secondary-btn";
  skip.textContent = "Skip";
  actions.appendChild(share);
  actions.appendChild(skip);
  el.appendChild(actions);

  const result = document.createElement("div");
  result.className = "ai-result";
  el.appendChild(result);

  const done = () => {
    actions.remove();
    input.disabled = true;
    state.blockCursor += 1;
    renderNextBlock();
  };

  skip.addEventListener("click", done);
  share.addEventListener("click", async () => {
    const answer = input.value.trim();
    if (!answer) return done();
    share.disabled = true;
    skip.disabled = true;
    result.innerHTML = `<div class="ai-thinking">…</div>`;
    try {
      const res = await callCoach({ mode: "reflection", blockId: block.id, answer });
      result.innerHTML = `<div class="coach-note">${escapeHtml(res.reply || "Thanks for sharing.")}</div>`;
    } catch (err) {
      result.innerHTML = `<div class="coach-note">Thanks for sharing. Keep that in mind as you go.</div>`;
    }
    done();
  });
}

function renderGamePlanBlock(el, block) {
  el.classList.add("block-ai");
  const sections = block.sections.map((s) => ({
    ...s,
    data: loadLocal(gamePlanKey(state.currentTrainingId, s.lessonId, s.blockId)),
  }));

  const sectionHtml = sections
    .map((s) => {
      if (!s.data) {
        return `<h3>${escapeHtml(s.title)}</h3><p class="muted">Not completed yet. Finish ${escapeHtml(s.from)} to add this section.</p>`;
      }
      const rows = Object.keys(s.data.fields)
        .map(
          (k) =>
            `<h4>${escapeHtml((s.data.labels && s.data.labels[k]) || k)}</h4><p>${escapeHtml(s.data.fields[k]).replace(/\n/g, "<br>")}</p>`
        )
        .join("");
      return `<h3>${escapeHtml(s.title)}</h3>${rows}`;
    })
    .join("");

  el.innerHTML = `<div class="block-ai__label">🏁 ${escapeHtml(block.title || "My Sales Game Plan")}</div>
    ${block.introHtml || ""}
    <div class="gameplan">${sectionHtml}</div>`;

  const docHtml = `<!DOCTYPE html><html><head><meta charset="utf-8"><title>${escapeHtml(block.title || "My Sales Game Plan")}</title>
    <style>body{font-family:Calibri,Arial,sans-serif;color:#1a1f27;max-width:760px;margin:32px auto;line-height:1.5}h1{color:#12233b}h3{color:#1f5fae;border-bottom:1px solid #dfe3e8;padding-bottom:4px;margin-top:28px}h4{margin:14px 0 4px}</style></head>
    <body><h1>${escapeHtml(block.title || "My Sales Game Plan")}</h1><p>${escapeHtml(block.docSubtitle || "OneAccord Sales Training")} · ${new Date().toLocaleDateString()}</p>${sectionHtml}</body></html>`;

  const actions = document.createElement("div");
  actions.className = "lesson-nav-actions";

  const dl = document.createElement("button");
  dl.className = "primary-btn";
  dl.textContent = "Download (Word)";
  dl.addEventListener("click", () => {
    const blob = new Blob([docHtml], { type: "application/msword" });
    const a = document.createElement("a");
    a.href = URL.createObjectURL(blob);
    a.download = block.fileName || "My-Sales-Game-Plan.doc";
    document.body.appendChild(a);
    a.click();
    a.remove();
  });

  const print = document.createElement("button");
  print.className = "secondary-btn";
  print.textContent = "Print / save as PDF";
  print.addEventListener("click", () => {
    const w = window.open("", "_blank");
    if (!w) return;
    w.document.write(docHtml);
    w.document.close();
    w.focus();
    w.print();
  });

  actions.appendChild(dl);
  actions.appendChild(print);
  el.appendChild(actions);
}

/* ---------------- Lesson completion ---------------- */

function renderLessonComplete() {
  const stream = getStream();
  const meta = state.lessonMeta;

  recordCompletion(state.currentTrainingId, state.currentLessonId);

  const el = document.createElement("div");
  el.className = "lesson-complete";
  el.innerHTML = `<h2>🎉 ${meta.completionTitle || "Lesson complete!"}</h2>
    <p><strong>Key concepts:</strong></p>
    <ul>${(meta.keyConcepts || []).map((c) => `<li>${c}</li>`).join("")}</ul>
    ${meta.nextTeaser ? `<p>${meta.nextTeaser}</p>` : ""}
    ${meta.resources && meta.resources.length ? `<p><strong>Keep going:</strong></p>${renderLinks(meta.resources)}` : ""}`;
  stream.appendChild(el);

  const actions = document.createElement("div");
  actions.className = "lesson-nav-actions";

  const reviewBtn = document.createElement("button");
  reviewBtn.className = "primary-btn";
  reviewBtn.textContent = "Review this lesson again";
  reviewBtn.addEventListener("click", () => {
    const training = state.index.trainings.find((t) => t.id === state.currentTrainingId);
    const lesson = training.lessons.find((l) => l.id === state.currentLessonId);
    loadLesson(training, lesson);
  });
  actions.appendChild(reviewBtn);

  if (meta.nextLessonId) {
    const nextBtn = document.createElement("button");
    nextBtn.className = "primary-btn";
    nextBtn.textContent = "Start next lesson";
    nextBtn.addEventListener("click", () => {
      const training = state.index.trainings.find((t) => t.id === state.currentTrainingId);
      const nextLesson = training.lessons.find((l) => l.id === meta.nextLessonId);
      if (nextLesson) loadLesson(training, nextLesson);
    });
    actions.appendChild(nextBtn);
  }

  stream.appendChild(actions);
  renderNav();
}

/* ---------------- Completion tracking ----------------
 * Writes go to the SharePoint "Training Completions" list via the
 * record-completion Netlify Function; state.completions (loaded from
 * get-completions in loadCompletions()) is the in-session cache that
 * isLessonComplete() reads for the nav checkmarks.
 */

function completionKey(trainingId, lessonId) {
  return `${trainingId}:${lessonId}`;
}

async function recordCompletion(trainingId, lessonId, score) {
  state.completions.add(completionKey(trainingId, lessonId));
  renderNav();

  if (!state.personEmail) return;
  try {
    const res = await fetch("/.netlify/functions/record-completion", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({
        trainingId,
        lessonId,
        personEmail: state.personEmail,
        score: typeof score === "number" ? score : undefined,
      }),
    });
    if (!res.ok) throw new Error(`record-completion returned ${res.status}`);
  } catch (err) {
    console.warn("Could not record completion to SharePoint", err);
  }
}

function isLessonComplete(trainingId, lessonId) {
  return state.completions.has(completionKey(trainingId, lessonId));
}
