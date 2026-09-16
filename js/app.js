/**
 * OneAccord Training Portal — shell engine
 * ------------------------------------------------
 * Renders a nav from a "trainings index" and steps a learner through a
 * training's content blocks one at a time, gating progress on hands-on
 * confirmations and answered questions.
 *
 * DATA SOURCE (current): static JSON files under /data — this is a stand-in
 * for the SharePoint index list + per-training content packages described
 * in the project plan. Swapping /data/trainings-index.json for a call to
 * the SharePoint "Trainings" list (via Microsoft Graph) is the only change
 * needed later; the rendering code below does not care where the JSON came
 * from.
 *
 * COMPLETION TRACKING (current): localStorage, per browser, as a placeholder
 * only. Phase 3 of the plan replaces this with a call to a Netlify Function
 * that writes a row to the SharePoint "Completions" list — see the
 * `recordCompletion()` function below, which is the single place that will
 * change.
 *
 * BLOCK TYPES SUPPORTED (the four the first pilot needs — see README):
 *   text      — static explanatory copy
 *   video     — an embedded video (SharePoint stream link or a file)
 *   activity  — a hands-on "find it now" step the learner confirms in person
 *   question  — a fixed-answer multiple-choice or true/false check
 */

const state = {
  index: null,
  currentTrainingId: null,
  currentLessonId: null,
  lessonBlocks: [],
  blockCursor: 0,
};

const navEl = document.getElementById("training-list");
const mainEl = document.getElementById("portal-main");

init();

async function init() {
  try {
    const res = await fetch("data/trainings-index.json");
    state.index = await res.json();
    renderNav();
  } catch (err) {
    navEl.innerHTML = `<div style="padding:0 20px;color:#f5b7b1;">Couldn't load the training list.</div>`;
    console.error(err);
  }
}

/* ---------------- Nav ---------------- */

function renderNav() {
  navEl.innerHTML = "";
  state.index.trainings.forEach((training) => {
    const heading = document.createElement("div");
    heading.className = "nav-training";
    heading.textContent = training.title;
    navEl.appendChild(heading);

    training.lessons.forEach((lesson) => {
      const item = document.createElement("div");
      const isPending = lesson.status === "pending";
      const isDone = isLessonComplete(training.id, lesson.id);
      item.className = "nav-lesson" + (isPending ? " is-pending" : "");
      item.innerHTML = `<span class="nav-lesson__check">${isDone ? "✅" : "⬜"}</span><span>${lesson.title}</span>`;
      if (!isPending) {
        item.addEventListener("click", () => loadLesson(training, lesson));
      } else {
        item.title = "Content coming soon";
      }
      navEl.appendChild(item);
    });
  });
}

function markActiveInNav(lessonId) {
  document.querySelectorAll(".nav-lesson").forEach((el) => el.classList.remove("is-active"));
  // Re-render is simplest given the small nav size; keeps active-state logic in one place.
}

/* ---------------- Lesson loading ---------------- */

async function loadLesson(training, lesson) {
  state.currentTrainingId = training.id;
  state.currentLessonId = lesson.id;
  state.blockCursor = 0;

  mainEl.innerHTML = `<div class="lesson-header">
      <h1>${lesson.title}</h1>
      <div class="lesson-progress">${training.title}</div>
    </div>
    <div id="block-stream"></div>`;

  const res = await fetch(`data/trainings/${training.id}/${lesson.id}.json`);
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
    ${meta.nextTeaser ? `<p>${meta.nextTeaser}</p>` : ""}`;
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

/* ---------------- Completion tracking (placeholder) ----------------
 * TODO(phase 3): replace this pair of functions with calls to the Netlify
 * Function that writes/reads the SharePoint "Completions" list, per the
 * project plan. Everything else in this file stays the same.
 */

function completionKey(trainingId, lessonId) {
  return `oa-training-complete:${trainingId}:${lessonId}`;
}

function recordCompletion(trainingId, lessonId) {
  try {
    localStorage.setItem(completionKey(trainingId, lessonId), new Date().toISOString());
  } catch (err) {
    console.warn("Could not record completion locally", err);
  }
}

function isLessonComplete(trainingId, lessonId) {
  try {
    return !!localStorage.getItem(completionKey(trainingId, lessonId));
  } catch (err) {
    return false;
  }
}
