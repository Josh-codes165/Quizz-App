const prevBtn = document.getElementById("previous-btn");
const nextBtn = document.getElementById("next-btn");
const input = document.getElementById("filestext");
const begin = document.getElementById("begin"); // if you don't use this, you can remove it

const QUIZ_DURATION = 120; // seconds
let remaining = QUIZ_DURATION;
let timerId = null;
let isFinished = false;
let timerStarted = false;

let questions = [];
let currentIndex = 0;

// ==== FILE UPLOAD & PARSE ====
input.addEventListener("change", async () => {
  const file = input.files[0];
  if (!file) {
    console.log("Pick a file first!");
    return;
  }

  const buf = await file.arrayBuffer();
  const wb = XLSX.read(buf, { type: "array" });
  const sheetName = wb.SheetNames[0];
  const ws = wb.Sheets[sheetName];
  const rows = XLSX.utils.sheet_to_json(ws, { defval: "" });

  console.log("Rows:", rows);
  questions = normalizeRows(rows);
  const problems = validateQuestions(questions);
  console.log(questions, problems);

  if (problems.length > 0) {
    alert(problems.join("\n"));
    return;
  }

  // Reset state BUT DO NOT START TIMER YET
  isFinished = false;
  timerStarted = false;
  currentIndex = 0;
  stopTimer();
  remaining = QUIZ_DURATION;
  updateTimer();

  render();
});

// Only keep question text, ignore options/answers
const normalizeRows = (rows) => {
  return rows.map((row, i) => {
    return {
      id: i + 1,
      text: String(row["Question"] || "").trim(),
    };
  });
};

function validateQuestions(qs) {
  const errors = [];
  qs.forEach((q, i) => {
    if (!q.text) errors.push(`Row ${i + 2}: "Question" is empty.`);
  });
  return errors;
}

// ==== RENDER ====
function render() {
  if (!questions.length) {
    alert("No questions loaded.");
    return;
  }

  const q = questions[currentIndex];
  const questionEl = document.getElementById("question");
  const optionEl = document.getElementById("answer-btn"); // we just reuse this as a container

  if (questionEl) {
    questionEl.textContent = q.text || `Question ${currentIndex + 1}`;
  }

  if (optionEl) {
    // No options, just show maybe a placeholder or keep it empty
    optionEl.innerHTML = ""; 
  }

  const progress = document.getElementById("quizCount");
  if (progress)
    progress.textContent = `Question ${currentIndex + 1} of ${questions.length}`;

  // Enable/disable navigation based on position / finished state
  if (prevBtn) prevBtn.disabled = isFinished || currentIndex === 0;
  if (nextBtn) nextBtn.disabled = isFinished || currentIndex === questions.length - 1;
}

// ==== NAVIGATION ====
prevBtn.addEventListener("click", () => {
  if (isFinished) return;
  if (!questions.length) return;

  currentIndex = Math.max(0, currentIndex - 1);
  render();
});

nextBtn.addEventListener("click", () => {
  if (isFinished) return;
  if (!questions.length) return;

  // 🔥 Start timer ONLY on first NEXT
  if (!timerStarted) {
    timerStarted = true;
    startTime();
  }

  currentIndex = Math.min(questions.length - 1, currentIndex + 1);
  render();
});

// ==== TIMER ====
function formatTime(sec) {
  const m = Math.floor(sec / 60).toString().padStart(2, "0");
  const s = (sec % 60).toString().padStart(2, "0");
  return `${m}:${s}`;
}

function updateTimer() {
  const el = document.getElementById("timer");
  if (!el) return;
  el.textContent = formatTime(remaining);
  el.classList.toggle("warning", remaining <= 10);
}

function startTime() {
  if (timerId) clearInterval(timerId);
  remaining = QUIZ_DURATION;
  updateTimer();

  timerId = setInterval(() => {
    if (isFinished) return;
    remaining -= 1;
    updateTimer();

    if (remaining <= 0) {
      finishQuiz("time");
    }
  }, 1000);
}

function stopTimer() {
  if (timerId) {
    clearInterval(timerId);
    timerId = null;
  }
}

// ==== FINISH / RESULT (no score, just time info) ====
function finishQuiz(reason = "manual") {
  if (isFinished) return;
  isFinished = true;

  stopTimer();

  if (prevBtn) prevBtn.disabled = true;
  if (nextBtn) nextBtn.disabled = true;

  const questionEl = document.getElementById("question");
  const optionEl = document.getElementById("answer-btn");
  const progress = document.getElementById("quizCount");

  const used = QUIZ_DURATION - Math.max(0, remaining);
  let title = "Quiz Completed";
  if (reason === "time") title = "Time's Up!";

  if (questionEl) questionEl.textContent = title;
  if (optionEl) {
    optionEl.innerHTML = `
      <p><strong>Time Used:</strong> ${formatTime(used)}</p>
    `;
  }
  if (progress) progress.textContent = "Done";
}
