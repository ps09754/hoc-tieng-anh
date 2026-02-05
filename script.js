const vocabBody = document.getElementById("vocabBody");
const topicTitle = document.getElementById("topicTitle");
const topicSubtitle = document.getElementById("topicSubtitle");
const downloadTemplateBtn = document.getElementById("downloadTemplateBtn");
const uploadExcelBtn = document.getElementById("uploadExcelBtn");
const resetCustomBtn = document.getElementById("resetCustomBtn");
const excelInput = document.getElementById("excelInput");
const autoPlayBtn = document.getElementById("autoPlayBtn");
const bilingualBtn = document.getElementById("bilingualBtn");
const stopBtn = document.getElementById("stopBtn");
const autoStatus = document.getElementById("autoStatus");
const volumeInput = document.getElementById("volume");
const volumeLabel = document.getElementById("volumeLabel");
const delayInput = document.getElementById("delayInput");
const rateInput = document.getElementById("rateInput");
const rateLabel = document.getElementById("rateLabel");
const startInput = document.getElementById("startInput");
const endInput = document.getElementById("endInput");
const loopInput = document.getElementById("loopInput");
const expandBtn = document.getElementById("expandBtn");
const pipBtn = document.getElementById("pipBtn");
const replayBtn = document.getElementById("replayBtn");
const fullOverlay = document.getElementById("fullOverlay");
const overlayWord = document.getElementById("overlayWord");
const overlayMeaning = document.getElementById("overlayMeaning");
const shrinkBtn = document.getElementById("shrinkBtn");
const overlayReplayBtn = document.getElementById("overlayReplayBtn");

// Pagination Elements
const prevPageBtn = document.getElementById("prevPageBtn");
const nextPageBtn = document.getElementById("nextPageBtn");
const pageInfo = document.getElementById("pageInfo");

let vocabList = [];
let defaultList = [];
let listSource = "default";
let hasCustomList = false;
let autoPlaying = false;
let currentIndex = -1;
let volume = Number(volumeInput.value);
let delaySeconds = Number(delayInput.value);
let rate = Number(rateInput.value);
let lastSpokenIndex = -1;
let replayPromise = null;
let bilingualMode = false;

// Pagination State
let currentPage = 1;
const itemsPerPage = 50;

rateLabel.textContent = `${rate.toFixed(2)}x`;

let pipCanvas = null;
let pipCtx = null;
let pipVideo = null;
let pipActive = false;

const storageKeys = {
  lastIndex: "vocab:lastIndex",
  rangeStart: "vocab:rangeStart",
  rangeEnd: "vocab:rangeEnd",
  loop: "vocab:loop",
  mode: "vocab:lastMode",
  customList: "vocab:customList",
};
const storageAvailable = (() => {
  try {
    const testKey = "__vocab_test__";
    window.localStorage.setItem(testKey, "1");
    window.localStorage.removeItem(testKey);
    return true;
  } catch (error) {
    return false;
  }
})();

const sleep = (ms) => new Promise((resolve) => setTimeout(resolve, ms));
const cleanText = (value) => {
  if (typeof value !== "string") return "";
  return value.trim().normalize("NFC");
};
const cleanTranscription = (value) => {
  const raw = cleanText(value);
  if (!raw) return "";
  return raw.replace(/^\/+|\/+$/g, "").trim();
};
const normalizeItem = (item) => {
  return {
    text: cleanText(item?.text),
    transcription: cleanTranscription(item?.transcription),
    vietnamese_meaning: cleanText(item?.vietnamese_meaning),
  };
};
const clampNumber = (value, min, max) => Math.min(Math.max(value, min), max);
const readStoredNumber = (key) => {
  if (!storageAvailable) return null;
  const raw = window.localStorage.getItem(key);
  if (raw == null) return null;
  const parsed = Number(raw);
  return Number.isNaN(parsed) ? null : parsed;
};
const readStoredBoolean = (key) => {
  if (!storageAvailable) return null;
  const raw = window.localStorage.getItem(key);
  if (raw == null) return null;
  return raw === "1";
};
const saveToStorage = (key, value) => {
  if (!storageAvailable) return;
  window.localStorage.setItem(key, String(value));
};
const getProgressKey = (source) => `${storageKeys.lastIndex}:${source}`;
const saveProgress = (index) => {
  if (index < 0) return;
  saveToStorage(getProgressKey(listSource), index);
};
const loadCustomListFromStorage = () => {
  if (!storageAvailable) return [];
  const raw = window.localStorage.getItem(storageKeys.customList);
  if (!raw) return [];
  try {
    const parsed = JSON.parse(raw);
    if (!Array.isArray(parsed)) return [];
    return parsed.map(normalizeItem).filter((item) => item.text);
  } catch (error) {
    return [];
  }
};
const setCustomState = (value) => {
  hasCustomList = value;
  resetCustomBtn.disabled = !hasCustomList;
};
const normalizeRangeInputs = () => {
  const total = vocabList.length || 1;
  let start = Number(startInput.value);
  let end = Number(endInput.value);
  if (Number.isNaN(start) || start < 1) start = 1;
  if (Number.isNaN(end) || end < 1) end = total;
  start = clampNumber(start, 1, total);
  end = clampNumber(end, 1, total);
  if (end < start) end = start;
  startInput.value = String(start);
  endInput.value = String(end);
  saveToStorage(storageKeys.rangeStart, start);
  saveToStorage(storageKeys.rangeEnd, end);
  return { startIndex: start - 1, endIndex: end - 1 };
};
const updateTopicInfo = () => {
  if (listSource === "custom") {
    topicTitle.textContent = "Chủ đề: Từ vựng của bạn";
    topicSubtitle.textContent = `Tổng số từ: ${vocabList.length} (Excel)`;
  } else {
    topicTitle.textContent = "Chủ đề: Tất cả";
    topicSubtitle.textContent = `Tổng số từ: ${vocabList.length}`;
  }
};
const applyVocabList = (list, source, options = {}) => {
  listSource = source;
  vocabList = list;
  currentPage = 1; // Reset to page 1
  updateTopicInfo();
  renderTableWithPagination(); // Updated render
  setCustomState(listSource === "custom");

  const total = vocabList.length || 1;
  startInput.max = String(total);
  endInput.max = String(total);

  if (options.resetRange) {
    startInput.value = "1";
    endInput.value = String(total);
    saveToStorage(storageKeys.rangeStart, 1);
    saveToStorage(storageKeys.rangeEnd, total);
  } else {
    const storedStart = readStoredNumber(storageKeys.rangeStart);
    const storedEnd = readStoredNumber(storageKeys.rangeEnd);

    // Heuristic: If stored range is 1-1 but list is larger, reset it to full.
    // Or if storedEnd is missing.
    if (!storedEnd || (storedStart === 1 && storedEnd === 1 && total > 1)) {
      startInput.value = "1";
      endInput.value = String(total);
      saveToStorage(storageKeys.rangeStart, 1);
      saveToStorage(storageKeys.rangeEnd, total);
    } else {
      if (storedStart != null) {
        startInput.value = String(clampNumber(storedStart, 1, total));
      } else {
        startInput.value = "1";
      }
      if (storedEnd != null) {
        endInput.value = String(clampNumber(storedEnd, 1, total));
      } else {
        endInput.value = String(total);
      }
    }
  }

  const storedLoop = readStoredBoolean(storageKeys.loop);
  if (storedLoop != null) {
    loopInput.checked = storedLoop;
  }

  const { startIndex, endIndex } = normalizeRangeInputs();
  let storedIndex = readStoredNumber(getProgressKey(listSource));
  // If stored index is valid, use it. Otherwise start at beginning of range.
  if (storedIndex != null && storedIndex >= 0 && storedIndex < total) {
    // Ensure stored index is within current range
    const safeIndex = clampNumber(storedIndex, startIndex, endIndex);
    currentIndex = safeIndex;
  } else {
    currentIndex = startIndex;
  }

  lastSpokenIndex = currentIndex;

  if (currentIndex >= 0 && vocabList[currentIndex]) {
    setActiveRow(currentIndex);
    updatePlayer(vocabList[currentIndex]);
  } else {
    setActiveRow(-1);
    updatePlayer(null);
  }
};
const normalizeHeader = (value) => {
  return cleanText(value)
    .toLowerCase()
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .replace(/\s+/g, " ");
};
const extractItemsFromRows = (rows) => {
  if (!rows.length) return [];
  const header = rows[0] || [];
  const normalized = header.map(normalizeHeader);
  const targetHeaders = {
    text: "tu vung",
    transcription: "phien am",
    meaning: "nghia tieng viet",
  };
  const indexes = {
    text: normalized.indexOf(targetHeaders.text),
    transcription: normalized.indexOf(targetHeaders.transcription),
    meaning: normalized.indexOf(targetHeaders.meaning),
  };
  const hasHeader =
    indexes.text !== -1 &&
    (indexes.transcription !== -1 || indexes.meaning !== -1);
  const dataRows = hasHeader ? rows.slice(1) : rows;
  const textIndex = hasHeader ? indexes.text : 0;
  const transcriptionIndex = hasHeader ? indexes.transcription : 1;
  const meaningIndex = hasHeader ? indexes.meaning : 2;

  return dataRows
    .map((row) => {
      return normalizeItem({
        text: row[textIndex],
        transcription: row[transcriptionIndex],
        vietnamese_meaning: row[meaningIndex],
      });
    })
    .filter((item) => item.text);
};
const saveCustomList = (items) => {
  if (!storageAvailable) return;
  window.localStorage.setItem(storageKeys.customList, JSON.stringify(items));
};
const clearCustomList = () => {
  if (!storageAvailable) return;
  window.localStorage.removeItem(storageKeys.customList);
  window.localStorage.removeItem(getProgressKey("custom"));
};
const applyCustomList = (items, options = {}) => {
  const normalizedItems = items.map(normalizeItem).filter((item) => item.text);
  if (!normalizedItems.length) return false;
  clearCustomList();
  saveCustomList(normalizedItems);
  applyVocabList(normalizedItems, "custom", { resetRange: true, ...options });
  saveProgress(currentIndex);
  return true;
};
const buildTemplateRows = () => {
  return [
    ["Từ vựng", "Phiên âm", "Nghĩa tiếng việt"],
    ["example", "ˈeg.zæm.pəl", "ví dụ"],
  ];
};
const downloadTemplate = () => {
  if (!window.XLSX) {
    topicSubtitle.textContent = "Không thể tải mẫu Excel. Thiếu thư viện XLSX.";
    return;
  }
  const worksheet = window.XLSX.utils.aoa_to_sheet(buildTemplateRows());
  const workbook = window.XLSX.utils.book_new();
  window.XLSX.utils.book_append_sheet(workbook, worksheet, "Vocabulary");
  window.XLSX.writeFile(workbook, "vocabulary-template.xlsx");
};
const handleExcelFile = async (file) => {
  if (!window.XLSX) {
    topicSubtitle.textContent =
      "Không thể đọc file Excel. Thiếu thư viện XLSX.";
    return;
  }
  const buffer = await file.arrayBuffer();
  const workbook = window.XLSX.read(buffer, { type: "array" });
  const sheetName = workbook.SheetNames[0];
  if (!sheetName) {
    topicSubtitle.textContent = "Không tìm thấy sheet trong file Excel.";
    return;
  }
  const worksheet = workbook.Sheets[sheetName];
  const rows = window.XLSX.utils.sheet_to_json(worksheet, {
    header: 1,
    defval: "",
  });
  const items = extractItemsFromRows(rows);
  if (!items.length) {
    topicSubtitle.textContent =
      "Không tìm thấy dữ liệu hợp lệ trong file. Hãy dùng mẫu Excel.";
    return;
  }
  const applied = applyCustomList(items);
  if (!applied) {
    topicSubtitle.textContent =
      "Không thể áp dụng dữ liệu. Vui lòng kiểm tra file.";
  }
};

// Pagination Logic
const getTotalPages = () => Math.ceil((vocabList.length || 1) / itemsPerPage);

const goToPage = (page) => {
  const totalPages = getTotalPages();
  const safePage = clampNumber(page, 1, totalPages);
  if (currentPage !== safePage) {
    currentPage = safePage;
    renderTableWithPagination();
  }
};

const nextPage = () => goToPage(currentPage + 1);
const prevPage = () => goToPage(currentPage - 1);

const updatePaginationControls = () => {
  const totalPages = getTotalPages();
  pageInfo.textContent = `Trang ${currentPage} / ${totalPages}`;
  prevPageBtn.disabled = currentPage <= 1;
  nextPageBtn.disabled = currentPage >= totalPages;
};

const renderTableWithPagination = () => {
  vocabBody.innerHTML = "";
  const total = vocabList.length;
  if (!total) {
    updatePaginationControls();
    return;
  }

  const start = (currentPage - 1) * itemsPerPage;
  const end = Math.min(start + itemsPerPage, total);
  const itemsToShow = vocabList.slice(start, end);

  itemsToShow.forEach((item, relativeIndex) => {
    const absoluteIndex = start + relativeIndex;
    const transcription = item.transcription ? `/${item.transcription}/` : "";
    const row = document.createElement("tr");

    // Add active class if this row is currently selected
    if (absoluteIndex === currentIndex) {
      row.classList.add("active");
    }

    const columns = [
      absoluteIndex + 1,
      item.text,
      transcription,
      item.vietnamese_meaning,
    ];
    columns.forEach((value) => {
      const cell = document.createElement("td");
      cell.textContent = value;
      row.appendChild(cell);
    });
    const actionsCell = document.createElement("td");
    const enButton = document.createElement("button");
    enButton.className = "btn";
    enButton.dataset.index = String(absoluteIndex);
    enButton.dataset.lang = "en";
    enButton.textContent = "Nghe EN";
    const viButton = document.createElement("button");
    viButton.className = "btn secondary";
    viButton.dataset.index = String(absoluteIndex);
    viButton.dataset.lang = "vi";
    viButton.textContent = "Nghe VI";
    actionsCell.appendChild(enButton);
    actionsCell.appendChild(viButton);
    row.appendChild(actionsCell);
    vocabBody.appendChild(row);
  });

  updatePaginationControls();
};

const setActiveRow = (index) => {
  // If index is negative (e.g. stop state), just remove active class from all rows
  if (index < 0) {
    const rows = vocabBody.querySelectorAll("tr");
    rows.forEach((row) => row.classList.remove("active"));
    return;
  }

  // Check if we need to switch pages
  const targetPage = Math.floor(index / itemsPerPage) + 1;
  if (targetPage !== currentPage) {
    currentPage = targetPage;
    renderTableWithPagination();
  }

  const rows = vocabBody.querySelectorAll("tr");
  rows.forEach((row) => {
    // Find the button in this row to check its data-index
    const btn = row.querySelector("button[data-index]");
    if (btn) {
      const rowIndex = Number(btn.dataset.index);
      if (rowIndex === index) {
        row.classList.add("active");
        row.scrollIntoView({
          behavior: "smooth",
          block: "center",
          inline: "nearest",
        });
      } else {
        row.classList.remove("active");
      }
    }
  });
};

const speakWord = (word, lang) => {
  return new Promise(async (resolve) => {
    if (!window.speechSynthesis) {
      resolve();
      return;
    }
    const utterance = new SpeechSynthesisUtterance(word);
    utterance.lang = lang || "en-US";
    utterance.rate = rate;
    utterance.pitch = 1;
    utterance.volume = volume;
    utterance.onend = () => resolve();
    utterance.onerror = () => resolve();
    window.speechSynthesis.cancel();
    window.speechSynthesis.speak(utterance);
  });
};

const queueReplay = async () => {
  if (lastSpokenIndex < 0 || !vocabList[lastSpokenIndex]) return;
  const task = async () => {
    if (window.speechSynthesis) {
      window.speechSynthesis.cancel();
    }
    currentIndex = lastSpokenIndex;
    setActiveRow(currentIndex);
    updatePlayer(vocabList[currentIndex]);
    await speakWord(vocabList[currentIndex].text, "en-US");
  };
  replayPromise = replayPromise ? replayPromise.then(task) : task();
  await replayPromise;
  replayPromise = null;
};

const stopAll = () => {
  autoPlaying = false;
  autoStatus.textContent = "Tự động: Tắt";
  autoPlayBtn.disabled = false;
  bilingualBtn.disabled = false;
  stopBtn.disabled = true;
  setActiveRow(-1);
  if (window.speechSynthesis) {
    window.speechSynthesis.cancel();
  }
};

const updatePlayer = (item) => {
  if (!item) {
    overlayWord.textContent = "---";
    overlayMeaning.textContent = "---";
    drawPiP(null);
    return;
  }
  overlayWord.textContent = item.text;

  // Update transcription
  const ovTransc = document.getElementById("overlayTranscription");
  if (ovTransc) {
    ovTransc.textContent = item.transcription ? `/${item.transcription}/` : "";
  }

  overlayMeaning.textContent = item.vietnamese_meaning;
  drawPiP(item);
};

const startAutoPlay = async (mode) => {
  if (!vocabList.length) return;
  autoPlaying = true;
  bilingualMode = mode === "bilingual";
  saveToStorage(storageKeys.mode, bilingualMode ? "bilingual" : "en");
  autoStatus.textContent = bilingualMode
    ? "Tự động: Song ngữ"
    : "Tự động: Đang phát";
  autoPlayBtn.disabled = true;
  bilingualBtn.disabled = true;
  stopBtn.disabled = false;

  const { startIndex, endIndex } = normalizeRangeInputs();
  let nextIndex =
    currentIndex >= startIndex && currentIndex <= endIndex
      ? currentIndex
      : startIndex;

  while (autoPlaying) {
    for (let i = nextIndex; i <= endIndex; i += 1) {
      if (!autoPlaying) break;
      if (replayPromise) {
        await replayPromise;
        replayPromise = null;
      }
      currentIndex = i;
      lastSpokenIndex = i;
      saveProgress(i);
      setActiveRow(i);
      updatePlayer(vocabList[i]);
      await speakWord(vocabList[i].text, "en-US");
      if (!autoPlaying) break;
      if (bilingualMode) {
        await sleep(750);
        if (!autoPlaying) break;
        await speakWord(vocabList[i].vietnamese_meaning, "vi-VN");
      }
      if (!autoPlaying) break;
      if (replayPromise) {
        await replayPromise;
        replayPromise = null;
      }
      await sleep(delaySeconds * 1000);
    }

    if (!autoPlaying) break;
    if (loopInput.checked) {
      nextIndex = startIndex;
      continue;
    }
    break;
  }

  if (autoPlaying) {
    autoPlaying = false;
    autoStatus.textContent = "Tự động: Hoàn tất";
    autoPlayBtn.disabled = false;
    bilingualBtn.disabled = false;
    stopBtn.disabled = true;
    setActiveRow(-1);
  }
};

const bindRowButtons = () => {
  vocabBody.addEventListener("click", async (event) => {
    const target = event.target;
    if (target.tagName !== "BUTTON") return;
    const index = Number(target.getAttribute("data-index"));
    if (Number.isNaN(index)) return;
    const lang = target.getAttribute("data-lang");
    const speakText =
      lang == "vi"
        ? vocabList[index].vietnamese_meaning
        : vocabList[index].text;
    stopAll();
    currentIndex = index;
    lastSpokenIndex = index;
    saveProgress(index);
    setActiveRow(index);
    updatePlayer(vocabList[index]);
    await speakWord(speakText, lang == "vi" ? "vi-VN" : "en-US");
  });
};

const wrapText = (ctx, text, maxWidth) => {
  if (!text) return [];
  const words = text.split(" ");
  const lines = [];
  let line = "";
  words.forEach((word) => {
    const testLine = line ? `${line} ${word}` : word;
    if (ctx.measureText(testLine).width <= maxWidth) {
      line = testLine;
    } else {
      if (line) lines.push(line);
      line = word;
    }
  });
  if (line) lines.push(line);
  return lines;
};

const initPiP = () => {
  if (pipCanvas) return;
  pipCanvas = document.createElement("canvas");
  pipCanvas.width = 640;
  pipCanvas.height = 360;
  pipCtx = pipCanvas.getContext("2d");
  pipVideo = document.createElement("video");
  pipVideo.muted = true;
  pipVideo.playsInline = true;
  pipVideo.srcObject = pipCanvas.captureStream(30);
  pipVideo.addEventListener("loadedmetadata", () => {
    pipVideo.play().catch(() => { });
  });
  pipVideo.addEventListener("enterpictureinpicture", () => {
    pipActive = true;
  });
  pipVideo.addEventListener("leavepictureinpicture", () => {
    pipActive = false;
  });
};

const drawPiP = (item) => {
  if (!pipCtx || !pipCanvas) return;
  const ctx = pipCtx;
  const width = pipCanvas.width;
  const height = pipCanvas.height;
  ctx.fillStyle = "#1e1b16";
  ctx.fillRect(0, 0, width, height);

  if (!item) {
    ctx.fillStyle = "#fff";
    ctx.font = "bold 28px Georgia, serif";
    ctx.textAlign = "center";
    ctx.textBaseline = "middle";
    ctx.fillText("No data", width / 2, height / 2);
    return;
  }

  const wordFont = "bold 60px Georgia, serif";
  const transcriptionFont = "32px Georgia, serif";
  const meaningFont = "40px Georgia, serif";
  ctx.textAlign = "center";
  ctx.textBaseline = "middle";

  ctx.font = wordFont;
  const wordLines = wrapText(ctx, item.text, width - 80);
  ctx.font = transcriptionFont;
  const transcription = item.transcription
    ? `/${item.transcription}/`
    : "";
  const transcriptionLines = wrapText(ctx, transcription, width - 120);
  ctx.font = meaningFont;
  const meaningLines = wrapText(
    ctx,
    item.vietnamese_meaning,
    width - 160,
  );

  const wordLineHeight = 60;
  const transcriptionLineHeight = 40;
  const meaningLineHeight = 48;
  const totalHeight =
    wordLines.length * wordLineHeight +
    transcriptionLines.length * transcriptionLineHeight +
    40 +
    meaningLines.length * meaningLineHeight;
  let y = (height - totalHeight) / 2 + wordLineHeight / 2;

  ctx.font = wordFont;
  ctx.fillStyle = "#fef6e4";
  wordLines.forEach((line) => {
    ctx.fillText(line, width / 2, y);
    y += wordLineHeight;
  });

  ctx.font = transcriptionFont;
  ctx.fillStyle = "#d9caa1";
  transcriptionLines.forEach((line) => {
    ctx.fillText(line, width / 2, y);
    y += transcriptionLineHeight;
  });

  y += 8;
  ctx.font = meaningFont;
  ctx.fillStyle = "#f7e1b5";
  meaningLines.forEach((line) => {
    ctx.fillText(line, width / 2, y);
    y += meaningLineHeight;
  });
};

const loadData = async () => {
  try {
    const response = await fetch("vocabulary.json");
    if (!response.ok) {
      throw new Error(`HTTP ${response.status}`);
    }
    const data = await response.json();
    const topics = Array.isArray(data)
      ? data
      : data && data.Topic
        ? [data]
        : [];
    const storedCustomList = loadCustomListFromStorage();
    if (!topics.length) {
      if (storedCustomList.length) {
        applyVocabList(storedCustomList, "custom");
        return;
      }
      topicSubtitle.textContent = "Không có dữ liệu từ vựng.";
      return;
    }

    defaultList = topics.reduce((acc, topic) => {
      const items = Array.isArray(topic.data) ? topic.data : [];
      return acc.concat(items.map(normalizeItem));
    }, []);
    defaultList = defaultList.filter((item) => item.text);
    if (storedCustomList.length) {
      applyVocabList(storedCustomList, "custom");
    } else {
      applyVocabList(defaultList, "default");
    }
  } catch (error) {
    if (window.location.protocol === "file:") {
      topicSubtitle.textContent =
        "Không thể tải dữ liệu từ vựng. Hãy mở trang qua máy chủ cục bộ.";
    } else {
      topicSubtitle.textContent = "Không thể tải dữ liệu từ vựng.";
    }
  }
};

downloadTemplateBtn.addEventListener("click", () => {
  downloadTemplate();
});

uploadExcelBtn.addEventListener("click", () => {
  excelInput.click();
});

excelInput.addEventListener("change", async (event) => {
  const file = event.target.files[0];
  if (!file) return;
  await handleExcelFile(file);
  event.target.value = "";
});

resetCustomBtn.addEventListener("click", () => {
  if (!hasCustomList) return;
  clearCustomList();
  if (defaultList.length) {
    applyVocabList(defaultList, "default");
  }
});

autoPlayBtn.addEventListener("click", () => {
  startAutoPlay("en");
});

bilingualBtn.addEventListener("click", () => {
  startAutoPlay("bilingual");
});

stopBtn.addEventListener("click", () => {
  stopAll();
});

volumeInput.addEventListener("input", (event) => {
  const value = Number(event.target.value);
  volume = value;
  volumeLabel.textContent = `${Math.round(value * 100)}%`;
});

startInput.addEventListener("change", () => {
  if (!vocabList.length) return;
  normalizeRangeInputs();
});

endInput.addEventListener("change", () => {
  if (!vocabList.length) return;
  normalizeRangeInputs();
});

loopInput.addEventListener("change", () => {
  saveToStorage(storageKeys.loop, loopInput.checked ? "1" : "0");
});

delayInput.addEventListener("input", (event) => {
  const value = Number(event.target.value);
  if (!Number.isNaN(value) && value > 0) {
    delaySeconds = value;
  }
});

rateInput.addEventListener("input", (event) => {
  const value = Number(event.target.value);
  if (!Number.isNaN(value)) {
    rate = value;
    rateLabel.textContent = `${value.toFixed(2)}x`;
  }
});

expandBtn.addEventListener("click", () => {
  fullOverlay.classList.add("active");
  fullOverlay.setAttribute("aria-hidden", "false");
});

shrinkBtn.addEventListener("click", () => {
  fullOverlay.classList.remove("active");
  fullOverlay.setAttribute("aria-hidden", "true");
});

replayBtn.addEventListener("click", async () => {
  await queueReplay();
});

overlayReplayBtn.addEventListener("click", async () => {
  await queueReplay();
});

pipBtn.addEventListener("click", async () => {
  const pipSupported =
    "pictureInPictureEnabled" in document &&
    document.pictureInPictureEnabled;
  if (!pipSupported) return;
  initPiP();
  const currentItem =
    currentIndex >= 0 ? vocabList[currentIndex] : vocabList[0];
  drawPiP(currentItem || null);
  try {
    await pipVideo.play();
    if (document.pictureInPictureElement) {
      await document.exitPictureInPicture();
    } else {
      await pipVideo.requestPictureInPicture();
    }
  } catch (error) {
    return;
  }
});

const pipSupported =
  "pictureInPictureEnabled" in document &&
  document.pictureInPictureEnabled;
if (!pipSupported) {
  pipBtn.disabled = true;
  pipBtn.textContent = "PiP không hỗ trợ";
}

// Bind events for pagination
prevPageBtn.addEventListener("click", prevPage);
nextPageBtn.addEventListener("click", nextPage);

bindRowButtons();
loadData();
