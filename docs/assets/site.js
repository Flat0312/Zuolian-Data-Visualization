(function () {
  function normalize(value) {
    return (value || "").toString().trim().toLowerCase();
  }

  function setupListFilters() {
    const inputs = document.querySelectorAll("[data-list-filter]");
    inputs.forEach((input) => {
      const targetId = input.getAttribute("data-list-filter");
      const list = document.getElementById(targetId);
      if (!list) {
        return;
      }
      const items = Array.from(list.querySelectorAll("[data-search]"));
      const countNode = document.querySelector(`[data-count-for="${targetId}"]`);
      const emptyNode = document.querySelector(`[data-empty-for="${targetId}"]`);

      function applyFilter() {
        const query = normalize(input.value);
        let visible = 0;
        items.forEach((item) => {
          const haystack = normalize(item.dataset.search);
          const match = !query || haystack.includes(query);
          item.hidden = !match;
          if (match) {
            visible += 1;
          }
        });
        if (countNode) {
          countNode.textContent = query
            ? `当前显示 ${visible} 条结果`
            : `共 ${items.length} 条记录`;
        }
        if (emptyNode) {
          emptyNode.hidden = visible !== 0;
        }
      }

      input.addEventListener("input", applyFilter);
      applyFilter();
    });
  }

  function escapeHtml(value) {
    return (value || "")
      .replaceAll("&", "&amp;")
      .replaceAll("<", "&lt;")
      .replaceAll(">", "&gt;")
      .replaceAll('"', "&quot;");
  }

  function scoreRecord(record, tokens) {
    const title = normalize(record.title);
    const subtitle = normalize(record.subtitle);
    const text = normalize(record.text);
    let score = 0;
    for (const token of tokens) {
      if (!text.includes(token) && !title.includes(token) && !subtitle.includes(token)) {
        return -1;
      }
      if (title === token) {
        score += 16;
      } else if (title.includes(token)) {
        score += 8;
      } else if (subtitle.includes(token)) {
        score += 4;
      } else {
        score += 2;
      }
    }
    return score;
  }

  function setupSearchApp() {
    const app = document.querySelector("[data-search-index]");
    if (!app) {
      return;
    }

    const input = app.querySelector("[data-search-input]");
    const meta = app.querySelector("[data-search-meta]");
    const resultsNode = app.querySelector("[data-search-results]");
    if (!input || !meta || !resultsNode) {
      return;
    }

    const params = new URLSearchParams(window.location.search);
    const initialQuery = params.get("q") || "";
    if (initialQuery) {
      input.value = initialQuery;
    }

    let records = [];
    let loaded = false;

    function renderPlaceholder(message) {
      resultsNode.innerHTML = `
        <article class="search-result search-result--placeholder">
          <h2>${escapeHtml(message)}</h2>
          <p>可以搜索人物、事件、关系类型、地点与证据摘录。</p>
        </article>
      `;
    }

    function renderResults() {
      const query = normalize(input.value);
      if (!loaded) {
        renderPlaceholder("搜索索引加载中");
        return;
      }

      if (!query) {
        meta.textContent = `搜索索引共 ${records.length} 条记录。`;
        renderPlaceholder("输入关键词后开始检索");
        const nextUrl = window.location.pathname;
        window.history.replaceState(null, "", nextUrl);
        return;
      }

      const tokens = query.split(/\s+/).filter(Boolean);
      const matches = records
        .map((record) => ({ ...record, score: scoreRecord(record, tokens) }))
        .filter((record) => record.score >= 0)
        .sort((left, right) => right.score - left.score || left.title.localeCompare(right.title, "zh-CN"))
        .slice(0, 60);

      meta.textContent = `关键词“${input.value}”匹配到 ${matches.length} 条结果。`;
      const nextUrl = `${window.location.pathname}?q=${encodeURIComponent(input.value)}`;
      window.history.replaceState(null, "", nextUrl);

      if (!matches.length) {
        renderPlaceholder("没有找到匹配结果");
        return;
      }

      resultsNode.innerHTML = matches
        .map(
          (record) => `
            <article class="search-result">
              <div class="search-result__type">${escapeHtml(record.type)}</div>
              <h2>${escapeHtml(record.title)}</h2>
              <p>${escapeHtml(record.subtitle)}</p>
              <a class="search-result__link" href="${escapeHtml(record.url)}">打开条目</a>
            </article>
          `
        )
        .join("");
    }

    input.addEventListener("input", renderResults);
    renderPlaceholder("搜索索引加载中");

    fetch(app.dataset.searchIndex)
      .then((response) => response.json())
      .then((data) => {
        records = Array.isArray(data) ? data : [];
        loaded = true;
        renderResults();
      })
      .catch(() => {
        meta.textContent = "搜索索引加载失败。";
        renderPlaceholder("无法加载搜索索引");
      });
  }

  document.addEventListener("DOMContentLoaded", function () {
    setupListFilters();
    setupSearchApp();
  });
})();
