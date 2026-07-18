import React, { useMemo, useRef, useState } from "react";
import { createRoot } from "react-dom/client";
import {
  UploadCloud,
  Image,
  X,
  CheckCircle2,
  AlertTriangle,
  ChevronDown,
  Download,
  ArrowRight,
  FileText,
  ShieldCheck,
  Sparkles,
  Plus,
} from "lucide-react";
import PptxGenJS from "pptxgenjs";
import "./styles.css";

const API_URL = import.meta.env.VITE_API_URL || "/api/analyze";

async function prepareImage(file) {
  const bitmap = await createImageBitmap(file);
  const scale = Math.min(1, 1800 / Math.max(bitmap.width, bitmap.height));
  const canvas = document.createElement("canvas");
  canvas.width = Math.round(bitmap.width * scale);
  canvas.height = Math.round(bitmap.height * scale);
  canvas.getContext("2d").drawImage(bitmap, 0, 0, canvas.width, canvas.height);
  bitmap.close();
  return canvas.toDataURL("image/jpeg", 0.82);
}

const sevOrder = { Critical: 0, Major: 1, Minor: 2 };
const sevLabel = { Critical: "Critical", Major: "Major", Minor: "Minor" };

function App() {
  const [files, setFiles] = useState([]);
  const [stage, setStage] = useState("upload");
  const [progress, setProgress] = useState(0);
  const [activeStep, setActiveStep] = useState("all");
  const [activeSeverity, setActiveSeverity] = useState("all");
  const [openId, setOpenId] = useState(1);
  const [findings, setFindings] = useState([]);
  const [score, setScore] = useState(100);
  const [error, setError] = useState("");
  const inputRef = useRef(null);

  const addFiles = (incoming) => {
    const next = [...incoming]
      .filter((f) => f.type.startsWith("image/"))
      .map((f, i) => ({
        file: f,
        url: URL.createObjectURL(f),
        name: f.name,
        id: `${Date.now()}-${i}`,
      }));
    setFiles((prev) => [...prev, ...next].slice(0, 8));
  };
  const analyze = async () => {
    if (!files.length) return;
    setError("");
    setStage("analyzing");
    setProgress(8);
    const timer = setInterval(
      () =>
        setProgress((p) => {
          return Math.min(92, p + Math.ceil(Math.random() * 5));
        }),
      500,
    );
    try {
      const images = await Promise.all(
        files.map(({ file }) => prepareImage(file)),
      );
      const response = await fetch(API_URL, {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ images }),
      });
      const data = await response.json().catch(() => ({}));
      if (!response.ok)
        throw new Error(data.error || `Ошибка сервера (${response.status})`);
      const normalized = (data.findings || []).map((finding, index) => ({
        ...finding,
        id: index + 1,
      }));
      setFindings(normalized);
      setScore(data.score ?? 100);
      setOpenId(normalized[0]?.id ?? null);
      setProgress(100);
      setTimeout(() => setStage("report"), 350);
    } catch (err) {
      setError(err.message || "Не удалось выполнить анализ");
      setStage("upload");
    } finally {
      clearInterval(timer);
    }
  };
  const filtered = useMemo(
    () =>
      findings
        .filter(
          (f) =>
            (activeStep === "all" || f.step === activeStep) &&
            (activeSeverity === "all" || f.severity === activeSeverity),
        )
        .sort((a, b) => sevOrder[a.severity] - sevOrder[b.severity]),
    [findings, activeStep, activeSeverity],
  );
  const counts = useMemo(
    () =>
      findings.reduce((a, f) => (a[f.severity]++, a), {
        Critical: 0,
        Major: 0,
        Minor: 0,
      }),
    [findings],
  );
  const blocked = counts.Critical + counts.Major > 0;

  async function downloadPptx() {
    const pptx = new PptxGenJS();
    pptx.layout = "LAYOUT_WIDE";
    pptx.author = "CX Review";
    pptx.subject = "UX/UI audit";
    pptx.title = "CX Review — UX/UI audit";
    pptx.theme = {
      headFontFace: "Arial",
      bodyFontFace: "Arial",
      lang: "ru-RU",
    };
    const addHeader = (slide, kicker, title) => {
      slide.background = { color: "F7F7F4" };
      slide.addText(kicker.toUpperCase(), {
        x: 0.65,
        y: 0.45,
        w: 4,
        h: 0.25,
        fontFace: "Arial",
        fontSize: 9,
        bold: true,
        color: "77776F",
        charSpacing: 1.4,
      });
      slide.addText(title, {
        x: 0.65,
        y: 0.82,
        w: 11.8,
        h: 0.55,
        fontSize: 25,
        bold: true,
        color: "171714",
        margin: 0,
      });
    };
    let s = pptx.addSlide();
    s.background = { color: "171714" };
    s.addText("CX REVIEW", {
      x: 0.7,
      y: 0.6,
      w: 3,
      h: 0.3,
      fontSize: 11,
      bold: true,
      color: "B6F36B",
      charSpacing: 2,
    });
    s.addText("UX/UI-аудит\nклиентского пути", {
      x: 0.7,
      y: 1.55,
      w: 7.2,
      h: 1.35,
      fontSize: 34,
      bold: true,
      color: "FFFFFF",
      breakLine: true,
      margin: 0,
    });
    s.addText(
      "Автоматизированный разбор экранов и рекомендации перед релизом",
      {
        x: 0.72,
        y: 3.15,
        w: 6.6,
        h: 0.6,
        fontSize: 15,
        color: "B8B8B2",
        margin: 0,
      },
    );
    s.addShape(pptx.ShapeType.roundRect, {
      x: 9.1,
      y: 1.45,
      w: 3.25,
      h: 2.3,
      rectRadius: 0.12,
      fill: { color: blocked ? "EF4B43" : "77C043" },
      line: { color: blocked ? "EF4B43" : "77C043" },
    });
    s.addText(blocked ? "НЕ ДОПУЩЕНО\nК РЕЛИЗУ" : "ДОПУЩЕНО\nК РЕЛИЗУ", {
      x: 9.45,
      y: 2.13,
      w: 2.55,
      h: 0.75,
      fontSize: 20,
      bold: true,
      color: "FFFFFF",
      align: "center",
      margin: 0,
    });
    s.addText(`${files.length} шага  •  ${findings.length} замечаний`, {
      x: 0.72,
      y: 6.55,
      w: 5,
      h: 0.25,
      fontSize: 11,
      color: "8A8A84",
      margin: 0,
    });
    s = pptx.addSlide();
    addHeader(
      s,
      "Результат анализа",
      blocked
        ? "Релиз заблокирован: есть Critical или Major"
        : "Блокирующих замечаний не обнаружено",
    );
    s.addText(String(counts.Critical), {
      x: 0.75,
      y: 1.75,
      w: 1.1,
      h: 0.65,
      fontSize: 31,
      bold: true,
      color: "EF4B43",
      margin: 0,
    });
    s.addText("Critical", {
      x: 0.75,
      y: 2.4,
      w: 1.3,
      h: 0.3,
      fontSize: 12,
      bold: true,
      color: "5C5C56",
      margin: 0,
    });
    s.addText(String(counts.Major), {
      x: 3.0,
      y: 1.75,
      w: 1.1,
      h: 0.65,
      fontSize: 31,
      bold: true,
      color: "E9822B",
      margin: 0,
    });
    s.addText("Major", {
      x: 3,
      y: 2.4,
      w: 1.3,
      h: 0.3,
      fontSize: 12,
      bold: true,
      color: "5C5C56",
      margin: 0,
    });
    s.addText(String(counts.Minor), {
      x: 5.25,
      y: 1.75,
      w: 1.1,
      h: 0.65,
      fontSize: 31,
      bold: true,
      color: "888880",
      margin: 0,
    });
    s.addText("Minor", {
      x: 5.25,
      y: 2.4,
      w: 1.3,
      h: 0.3,
      fontSize: 12,
      bold: true,
      color: "5C5C56",
      margin: 0,
    });
    s.addText("Что делать перед релизом", {
      x: 0.75,
      y: 3.5,
      w: 4,
      h: 0.3,
      fontSize: 16,
      bold: true,
      color: "171714",
      margin: 0,
    });
    s.addText(
      blocked
        ? "Исправить все Critical и Major замечания, пройти полный сценарий и запустить анализ повторно."
        : "Можно выпускать. Minor замечания допустимо включить в план последующих улучшений.",
      {
        x: 0.75,
        y: 3.95,
        w: 7.2,
        h: 1,
        fontSize: 15,
        color: "4B4B46",
        breakLine: true,
        margin: 0,
      },
    );
    findings.forEach((f, i) => {
      s = pptx.addSlide();
      addHeader(s, `Шаг ${f.step} · ${f.area}`, f.title);
      const c =
        f.severity === "Critical"
          ? "EF4B43"
          : f.severity === "Major"
            ? "E9822B"
            : "888880";
      s.addShape(pptx.ShapeType.roundRect, {
        x: 0.68,
        y: 1.55,
        w: 1.35,
        h: 0.4,
        rectRadius: 0.05,
        fill: { color: c },
        line: { color: c },
      });
      s.addText(f.severity, {
        x: 0.78,
        y: 1.65,
        w: 1.15,
        h: 0.16,
        fontSize: 10,
        bold: true,
        color: "FFFFFF",
        align: "center",
        margin: 0,
      });
      s.addText("Проблема", {
        x: 0.72,
        y: 2.35,
        w: 2,
        h: 0.3,
        fontSize: 12,
        bold: true,
        color: "77776F",
        margin: 0,
      });
      s.addText(f.detail, {
        x: 0.72,
        y: 2.75,
        w: 5.7,
        h: 1.15,
        fontSize: 16,
        color: "171714",
        breakLine: true,
        margin: 0,
      });
      s.addText("Рекомендация", {
        x: 6.95,
        y: 2.35,
        w: 2,
        h: 0.3,
        fontSize: 12,
        bold: true,
        color: "77776F",
        margin: 0,
      });
      s.addText(f.recommendation, {
        x: 6.95,
        y: 2.75,
        w: 5.5,
        h: 1.15,
        fontSize: 16,
        color: "171714",
        breakLine: true,
        margin: 0,
      });
      s.addText(
        `${String(i + 1).padStart(2, "0")} / ${String(findings.length).padStart(2, "0")}`,
        {
          x: 11.5,
          y: 6.8,
          w: 0.9,
          h: 0.2,
          fontSize: 9,
          color: "999990",
          align: "right",
          margin: 0,
        },
      );
    });
    await pptx.writeFile({ fileName: "CX-Review-report.pptx" });
  }

  if (stage === "analyzing")
    return <Analyzing progress={progress} count={files.length} />;
  if (stage === "report")
    return (
      <Report
        files={files}
        allFindings={findings}
        findings={filtered}
        counts={counts}
        blocked={blocked}
        score={score}
        activeStep={activeStep}
        setActiveStep={setActiveStep}
        activeSeverity={activeSeverity}
        setActiveSeverity={setActiveSeverity}
        openId={openId}
        setOpenId={setOpenId}
        download={downloadPptx}
        reset={() => {
          setStage("upload");
          setFiles([]);
          setFindings([]);
        }}
      />
    );
  return (
    <Upload
      files={files}
      addFiles={addFiles}
      setFiles={setFiles}
      inputRef={inputRef}
      analyze={analyze}
      error={error}
    />
  );
}

function Brand() {
  return (
    <div className="brand">
      <span className="brand-mark">CX</span>
      <span>REVIEW</span>
    </div>
  );
}
function Upload({ files, addFiles, setFiles, inputRef, analyze, error }) {
  const drop = (e) => {
    e.preventDefault();
    addFiles(e.dataTransfer.files);
  };
  return (
    <main className="shell upload-page">
      <header>
        <Brand />
        <span className="header-note">AI-аудит клиентского опыта</span>
      </header>
      <section className="hero">
        <div className="eyebrow">
          <Sparkles size={14} /> UX/UI QUALITY GATE
        </div>
        <h1>
          Найдём, что мешает
          <br />
          клиенту двигаться дальше.
        </h1>
        <p>
          Загрузите экраны клиентского пути. Мы разберём каждый шаг, оценим
          риски и соберём готовый отчёт для команды.
        </p>
      </section>
      <section className="upload-card">
        <div className="section-top">
          <div>
            <span className="step-no">01</span>
            <h2>Добавьте экраны</h2>
          </div>
          <span className="limit">до 8 файлов · PNG, JPG</span>
        </div>
        <div
          className={"dropzone " + (files.length ? "has-files" : "")}
          onDragOver={(e) => e.preventDefault()}
          onDrop={drop}
          onClick={() => inputRef.current.click()}
        >
          {files.length ? (
            <div className="thumbs">
              {files.map((f, i) => (
                <div
                  className="thumb"
                  key={f.id}
                  onClick={(e) => e.stopPropagation()}
                >
                  <img src={f.url} />
                  <span>{String(i + 1).padStart(2, "0")}</span>
                  <button
                    onClick={() =>
                      setFiles((x) => x.filter((v) => v.id !== f.id))
                    }
                  >
                    <X size={14} />
                  </button>
                </div>
              ))}
              {files.length < 8 && (
                <div
                  className="add-thumb"
                  onClick={() => inputRef.current.click()}
                >
                  <Plus />
                  <small>Добавить</small>
                </div>
              )}
            </div>
          ) : (
            <>
              <div className="upload-icon">
                <UploadCloud />
              </div>
              <strong>Перетащите изображения сюда</strong>
              <span>или нажмите, чтобы выбрать файлы</span>
            </>
          )}
          <input
            ref={inputRef}
            type="file"
            accept="image/*"
            multiple
            hidden
            onChange={(e) => addFiles(e.target.files)}
          />
        </div>
        <div className="flow-hint">
          <Image size={17} />
          <span>
            Расположите экраны в порядке прохождения сценария — слева направо.
          </span>
        </div>
        {error && (
          <div className="api-error">
            <AlertTriangle size={16} />
            <span>{error}</span>
          </div>
        )}
        <button className="primary" disabled={!files.length} onClick={analyze}>
          Запустить анализ <ArrowRight size={18} />
        </button>
      </section>
      <footer>
        <span>Privacy by design</span>
        <span>Изображения отправляются в OpenAI API только для анализа</span>
      </footer>
    </main>
  );
}

function Analyzing({ progress, count }) {
  const labels = [
    "Распознаём структуру интерфейса",
    "Восстанавливаем клиентский путь",
    "Проверяем UX-эвристики",
    "Оцениваем критичность",
    "Формируем рекомендации",
  ];
  const idx = Math.min(labels.length - 1, Math.floor(progress / 21));
  return (
    <main className="analyzing">
      <div className="analyze-wrap">
        <Brand />
        <div className="orbit">
          <div className="orbit-inner">
            <Sparkles size={34} />
          </div>
        </div>
        <span className="mono">АНАЛИЗ {count} ЭКРАНОВ</span>
        <h1>{labels[idx]}</h1>
        <p>Проверяем интерфейс, логику переходов и качество обратной связи.</p>
        <div className="progress">
          <i style={{ width: `${progress}%` }} />
        </div>
        <div className="progress-meta">
          <span>{progress}%</span>
          <span>Обычно это занимает меньше минуты</span>
        </div>
      </div>
    </main>
  );
}

function Report({
  files,
  allFindings,
  findings,
  counts,
  blocked,
  score,
  activeStep,
  setActiveStep,
  activeSeverity,
  setActiveSeverity,
  openId,
  setOpenId,
  download,
  reset,
}) {
  return (
    <main className="report">
      <aside>
        <Brand />
        <nav>
          <a className="active">
            <FileText />
            Отчёт
          </a>
          <a>
            <ShieldCheck />
            Методология
          </a>
        </nav>
        <div className="aside-bottom">
          <small>ПРОЕКТ</small>
          <strong>Новый клиентский путь</strong>
          <span>{files.length} экранов · сегодня</span>
          <button onClick={reset}>Начать заново</button>
        </div>
      </aside>
      <section className="report-main">
        <header>
          <div>
            <div className="crumb">ПРОЕКТ / РЕЗУЛЬТАТ АНАЛИЗА</div>
            <h1>Отчёт по клиентскому пути</h1>
          </div>
          <button className="download" onClick={download}>
            <Download size={17} /> Скачать PPTX
          </button>
        </header>
        <div className={"verdict " + (blocked ? "blocked" : "passed")}>
          <div className="verdict-icon">
            {blocked ? <AlertTriangle /> : <CheckCircle2 />}
          </div>
          <div>
            <span>ИТОГОВЫЙ ВЕРДИКТ</span>
            <h2>{blocked ? "Не допущено к релизу" : "Допущено к релизу"}</h2>
            <p>
              {blocked
                ? "Исправьте Critical и Major замечания и запустите проверку повторно."
                : "Блокирующих замечаний не обнаружено."}
            </p>
          </div>
          <div className="score">
            <strong>{score}</strong>
            <span>/ 100</span>
            <small>CX score</small>
          </div>
        </div>
        <div className="metrics">
          <Metric
            n={counts.Critical}
            label="Critical"
            tone="red"
            text="Блокирует релиз"
          />
          <Metric
            n={counts.Major}
            label="Major"
            tone="orange"
            text="Блокирует релиз"
          />
          <Metric
            n={counts.Minor}
            label="Minor"
            tone="gray"
            text="Можно исправить позже"
          />
          <div className="metric total">
            <strong>{counts.Critical + counts.Major + counts.Minor}</strong>
            <span>Всего замечаний</span>
            <small>в {files.length} шагах пути</small>
          </div>
        </div>
        <div className="content-grid">
          <div className="journey">
            <div className="block-title">
              <div>
                <span className="step-no">01</span>
                <h3>Клиентский путь</h3>
              </div>
              <span>{files.length} шага</span>
            </div>
            <div className="journey-strip">
              {files.map((f, i) => (
                <button
                  key={f.id}
                  className={activeStep === i + 1 ? "selected" : ""}
                  onClick={() =>
                    setActiveStep(activeStep === i + 1 ? "all" : i + 1)
                  }
                >
                  <div>
                    <img src={f.url} />
                    {allFindings.some((x) => x.step === i + 1) && (
                      <i
                        className={
                          allFindings.some(
                            (x) =>
                              x.step === i + 1 && x.severity === "Critical",
                          )
                            ? "crit"
                            : "maj"
                        }
                      />
                    )}
                  </div>
                  <span>Шаг {i + 1}</span>
                </button>
              ))}
            </div>
          </div>
          <div className="findings">
            <div className="block-title">
              <div>
                <span className="step-no">02</span>
                <h3>Замечания</h3>
              </div>
              <div className="filters">
                {["all", "Critical", "Major", "Minor"].map((x) => (
                  <button
                    className={activeSeverity === x ? "on" : ""}
                    onClick={() => setActiveSeverity(x)}
                    key={x}
                  >
                    {x === "all" ? "Все" : x}
                    {x !== "all" && ` ${counts[x]}`}
                  </button>
                ))}
              </div>
            </div>
            <div className="finding-list">
              {findings.map((f) => (
                <article
                  className={"finding " + f.severity.toLowerCase()}
                  key={f.id}
                >
                  <button
                    className="finding-head"
                    onClick={() => setOpenId(openId === f.id ? null : f.id)}
                  >
                    <span className="sev-dot" />
                    <div>
                      <small>
                        ШАГ {f.step} · {f.area.toUpperCase()}
                      </small>
                      <h4>{f.title}</h4>
                    </div>
                    <span className="sev-pill">{sevLabel[f.severity]}</span>
                    <ChevronDown className={openId === f.id ? "rot" : ""} />
                  </button>
                  {openId === f.id && (
                    <div className="finding-body">
                      <div className="evidence">
                        <small>ЧТО ВИДНО НА ЭКРАНЕ</small>
                        <p>{f.evidence}</p>
                      </div>
                      <div>
                        <small>ПОЧЕМУ ЭТО ПРОБЛЕМА</small>
                        <p>{f.detail}</p>
                      </div>
                      <div className="recommend">
                        <small>РЕКОМЕНДАЦИЯ</small>
                        <p>{f.recommendation}</p>
                      </div>
                    </div>
                  )}
                </article>
              ))}
              {!findings.length && (
                <div className="empty">
                  По выбранным фильтрам замечаний нет.
                </div>
              )}
            </div>
          </div>
        </div>
      </section>
    </main>
  );
}
function Metric({ n, label, tone, text }) {
  return (
    <div className={"metric " + tone}>
      <div>
        <i />
        <span>{label}</span>
      </div>
      <strong>{n}</strong>
      <small>{text}</small>
    </div>
  );
}

createRoot(document.getElementById("root")).render(<App />);
