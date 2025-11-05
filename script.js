const prevBtn = document.getElementById("previous-btn")
const nextBtn = document.getElementById("next-btn")
const input = document.getElementById("filestext");
const begin = document.getElementById("begin")
const  QUIZ_DURATION = 120;
const AUTO_ADVANCED_DELAY = 200;
let remaining = QUIZ_DURATION;
let timerId = null;
let autoAdvanced = null;
let isFinished = false;
let isAdvancing = false;
let questions = [];
let currentIndex = 0;
const answers = new Map();

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

  if(problems.length === 0){
  isFinished = false;
  answers.clear();
  currentIndex = 0;
  startTime();
  render();
}
  if (problems.length === 0) {
    render();
  }
});

const normalizeRows = (rows) => {
  const LETTERS = ["A", "B", "C", "D"];
  return rows.map((row, i) => {
    const options = [
      String(row["Option A"] ?? ""),
      String(row["Option B"] ?? ""),
      String(row["Option C"] ?? ""),
      String(row["Option D"] ?? "")
    ];
    const correctRaw = String(row["Answer"] ?? row["Correct"] ?? "").trim();
    let correctIndex = -1;

    const letter = correctRaw.toUpperCase();
    if (LETTERS.includes(letter)) {
      correctIndex = LETTERS.indexOf(letter);
    } else if (correctRaw) {
      const idx = options.findIndex(
        (o) => o.trim().toLowerCase() === correctRaw.toLowerCase()
      );
      if (idx >= 0) correctIndex = idx;
    }

    return {
      id: i + 1,
      text: String(row["Question"] || "").trim(),
      options,
      correctIndex,
    };
  });
};

function validateQuestions(qs) {
  const errors = [];
  qs.forEach((q, i) => {
    if (!q.text) errors.push(`Row ${i + 2}: "Question" is empty.`);
    if (!Array.isArray(q.options))
      errors.push(`Row ${i + 2}: "Options" missing or invalid.`);
    else if (q.options.filter(Boolean).length < 2)
      errors.push(`Row ${i + 2}: Need at least two options.`);
    if (q.correctIndex < 0)
      errors.push(`Row ${i + 2}: "Correct" must be A–D or match an option`);
  });
  return errors;
}

function render() {
  if (!questions.length) {
    console.warn("No questions to display!");
    return;
  }

  const q = questions[currentIndex];
  const questionEl = document.getElementById("question");
  const optionEl = document.getElementById("answer-btn");

  questionEl.textContent = q.text;
  optionEl.innerHTML = "";

  q.options.forEach((optText, idx) => {
    if (!optText) return;
    const li = document.createElement("li");
    const btn = document.createElement("button");
    btn.type = "button";
    btn.textContent = optText;
    btn.classList.add("btn")
    if (answers.get(q.id) === idx) btn.classList.add("selected");
    btn.addEventListener("click", () => {
      if(isFinished) return;
      if (isAdvancing) return;
      isAdvancing = true;
      answers.set(q.id, idx);
      const allButtons = optionEl.querySelectorAll("button");
      allButtons.forEach(b => b.disabled = true);
      btn.classList.add("selected")
      if(autoAdvanced) clearTimeout(autoAdvanced);
      autoAdvanced = setTimeout(goToNext, AUTO_ADVANCED_DELAY)
      render();
    });
    li.appendChild(btn);
    optionEl.appendChild(li);
  });

  const progress = document.getElementById("quizCount");
  if (progress)
    progress.textContent = `Question ${currentIndex + 1} of ${questions.length}`;
}

prevBtn.addEventListener("click", ()=>{
   currentIndex = Math.max(0, currentIndex -1);
   render()
});

nextBtn.addEventListener("click", ()=> {
   currentIndex = Math.min(questions.length -1, currentIndex + 1);
   render();
})

function goToNext(){
  if(isFinished) return;
   isAdvancing = false;
   if(currentIndex < questions.length - 1){
      currentIndex +=1;
      render();
   }else{
      finishQuiz()
   }
}

function showResult(){
   const total = questions.length;
   let score = 0;
   for(const q of questions){
      if(answers.get(q.id) === q.correctIndex) score +=1;
   }
   const questionEl = document.getElementById("question")
   const optionEl = document.getElementById("answer-btn")
   const progress = document.getElementById("quizCount")

   questionEl.textContent = "Quiz Completed";
   optionEl.innerHTML = `<p>Your Score; ${score}/ ${total}</p>`;
   if (progress) progress.textContent = "Done"
}

function formatTime(sec){
   const m = Math.floor(sec / 60).toString().padStart(2, '0');
   const s = (sec % 60).toString().padStart(2, "0");
   return `${m}:${s}`;
}

function updateTimer(){
   const el = document.getElementById("timer")
   if(!el) return;
   el.textContent = formatTime(remaining);
   el.classList.toggle("warning", remaining <= 10);
}

function startTime(){
   if(timerId) clearInterval(timerId);
   remaining = QUIZ_DURATION;
   updateTimer();

   timerId = setInterval(() => {
      if(isFinished) return;
      remaining -= 1
      updateTimer()

      if(remaining <= 0){
         finishQuiz("time")
      }
   }, 1000);
}

function stopTimer(){
   if(timerId){
      clearInterval(timerId);
      timerId = null;
   }
}

function finishQuiz(reason = "manual") {
   if(isFinished) return;
   isFinished = true;

   stopTimer();
   if(autoAdvanced) clearTimeout(autoAdvanced);

   if(prevBtn) prevBtn.disabled = true;
   if(nextBtn) nextBtn.disabled = true;

   const optionEl = document.getElementById("answer-btn");
   if(optionEl){
      optionEl.querySelectorAll("button").forEach(b => (b.disabled = true));
   }
   showResult(reason);
}

function showResult(reason ="manual"){
  const total = questions.length;
  let score = 0;
  for (const q of questions){
    if(answers.get(q.id) === q.correctIndex) score +=1;
  }

  const questionEl = document.getElementById("question");
  const optionEl = document.getElementById("answer-btn")
  const progress = document.getElementById("quizCount")

  const used = QUIZ_DURATION - Math.max(0, remaining);

  let title = "Quiz Completed"
  if(reason === "time") title = "Time's Up!";

  if(questionEl) questionEl.textContent = title;
  if(optionEl){
    optionEl.innerHTML = `<p><strong>Your Score:</strong> ${score} / ${total} </p>
    <p><strong>Time Used:</strong> ${formatTime(used)}</p>`;
  }
  
  if(progress) progress.textContent = "Done"
}

