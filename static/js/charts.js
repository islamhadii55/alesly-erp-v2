(function () {
  "use strict";

  var NS = "http://www.w3.org/2000/svg";

  function svgEl(name, attrs) {
    var node = document.createElementNS(NS, name);
    if (attrs) {
      for (var key in attrs) {
        if (Object.prototype.hasOwnProperty.call(attrs, key)) {
          node.setAttribute(key, attrs[key]);
        }
      }
    }
    return node;
  }

  function fmt(value) {
    var n = Number(value) || 0;
    return n.toLocaleString("ar-EG", { maximumFractionDigits: 2 });
  }

  function shortNum(value) {
    var n = Number(value) || 0;
    var abs = Math.abs(n);
    if (abs >= 1000000) return (n / 1000000).toFixed(1) + "م";
    if (abs >= 1000) return (n / 1000).toFixed(1) + "ألف";
    return n.toFixed(0);
  }

  function niceMax(value) {
    if (value <= 0) return 10;
    var exp = Math.pow(10, Math.floor(Math.log10(value)));
    var frac = value / exp;
    var nice = frac <= 1 ? 1 : frac <= 2 ? 2 : frac <= 2.5 ? 2.5 : frac <= 5 ? 5 : 10;
    return nice * exp;
  }

  function makeRoot(container) {
    var box = container.getBoundingClientRect();
    var width = Math.max(320, box.width || container.clientWidth || 640);
    var height = Number(container.dataset.height || 300);
    var svg = svgEl("svg", {
      viewBox: "0 0 " + width + " " + height,
      width: "100%",
      height: height,
      preserveAspectRatio: "xMidYMid meet",
      class: "chart-svg",
      dir: "ltr",
    });
    container.innerHTML = "";
    container.appendChild(svg);
    return { svg: svg, width: width, height: height };
  }

  function addDefs(svg, id, colors) {
    var defs = svgEl("defs");
    colors.forEach(function (color, index) {
      var grad = svgEl("linearGradient", { id: id + "-" + index, x1: "0", y1: "0", x2: "0", y2: "1" });
      grad.appendChild(svgEl("stop", { offset: "0%", "stop-color": color, "stop-opacity": "0.28" }));
      grad.appendChild(svgEl("stop", { offset: "100%", "stop-color": color, "stop-opacity": "0" }));
      defs.appendChild(grad);
    });
    svg.appendChild(defs);
  }

  function smoothPath(points) {
    if (!points.length) return "";
    if (points.length === 1) return "M" + points[0][0] + "," + points[0][1];
    var d = "M" + points[0][0] + "," + points[0][1];
    for (var i = 0; i < points.length - 1; i++) {
      var p0 = points[i === 0 ? 0 : i - 1];
      var p1 = points[i];
      var p2 = points[i + 1];
      var p3 = points[i + 2 >= points.length ? points.length - 1 : i + 2];
      var c1x = p1[0] + (p2[0] - p0[0]) / 6;
      var c1y = p1[1] + (p2[1] - p0[1]) / 6;
      var c2x = p2[0] - (p3[0] - p1[0]) / 6;
      var c2y = p2[1] - (p3[1] - p1[1]) / 6;
      d += " C" + c1x + "," + c1y + " " + c2x + "," + c2y + " " + p2[0] + "," + p2[1];
    }
    return d;
  }

  function drawLegend(container, series) {
    var legend = document.createElement("div");
    legend.className = "chart-legend";
    series.forEach(function (item) {
      var span = document.createElement("span");
      span.className = "chart-legend-item";
      var dot = document.createElement("i");
      dot.style.background = item.color;
      span.appendChild(dot);
      span.appendChild(document.createTextNode(item.name));
      legend.appendChild(span);
    });
    container.appendChild(legend);
  }

  function lineChart(container, data) {
    var series = data.series || [];
    if (!series.length || !data.labels || !data.labels.length) {
      container.innerHTML = '<p class="chart-empty">لا توجد بيانات كافية لعرض الرسم</p>';
      return;
    }
    var root = makeRoot(container);
    var width = root.width;
    var height = root.height;
    var padTop = 24;
    var padBottom = 42;
    var padSide = 62;
    var plotW = width - padSide * 2;
    var plotH = height - padTop - padBottom;
    var count = data.labels.length;

    var max = 0;
    series.forEach(function (s) {
      (s.values || []).forEach(function (v) {
        max = Math.max(max, Number(v) || 0);
      });
    });
    max = niceMax(max);

    addDefs(root.svg, "line-grad", series.map(function (s) { return s.color; }));

    var gridCount = 4;
    for (var g = 0; g <= gridCount; g++) {
      var y = padTop + (plotH / gridCount) * g;
      root.svg.appendChild(svgEl("line", {
        x1: padSide, y1: y, x2: width - padSide, y2: y,
        stroke: "#eef2f7", "stroke-width": g === gridCount ? 1.4 : 1,
      }));
      var label = svgEl("text", {
        x: padSide - 10, y: y + 4, "text-anchor": "end",
        class: "chart-axis", "font-size": "11",
      });
      label.textContent = shortNum(max - (max / gridCount) * g);
      root.svg.appendChild(label);
    }

    var stepX = count > 1 ? plotW / (count - 1) : 0;
    var skip = Math.ceil(count / 12);
    data.labels.forEach(function (name, i) {
      if (i % skip !== 0 && i !== count - 1) return;
      var x = padSide + stepX * i;
      var label = svgEl("text", {
        x: x, y: height - padBottom + 22, "text-anchor": "middle",
        class: "chart-axis", "font-size": "11",
      });
      label.textContent = name;
      root.svg.appendChild(label);
    });

    series.forEach(function (item, index) {
      var pts = (item.values || []).map(function (v, i) {
        var x = count > 1 ? padSide + stepX * i : padSide + plotW / 2;
        var y = padTop + plotH - (plotH * (Number(v) || 0)) / max;
        return [x, y];
      });
      var path = smoothPath(pts);
      if (item.area !== false) {
        var areaPath = path + " L" + pts[pts.length - 1][0] + "," + (padTop + plotH) +
          " L" + pts[0][0] + "," + (padTop + plotH) + " Z";
        root.svg.appendChild(svgEl("path", { d: areaPath, fill: "url(#line-grad-" + index + ")" }));
      }
      root.svg.appendChild(svgEl("path", {
        d: path, fill: "none", stroke: item.color, "stroke-width": 2.6,
        "stroke-linecap": "round", "stroke-linejoin": "round",
      }));
      pts.forEach(function (p, i) {
        var circle = svgEl("circle", {
          cx: p[0], cy: p[1], r: data.labels.length > 16 ? 2.4 : 4,
          fill: "#fff", stroke: item.color, "stroke-width": 2.2,
        });
        var title = svgEl("title");
        title.textContent = (item.name || "") + ": " + fmt((item.values || [])[i]) + " · " + data.labels[i];
        circle.appendChild(title);
        root.svg.appendChild(circle);
      });
    });
    drawLegend(container, series);
  }

  function barChart(container, data) {
    var labels = data.labels || [];
    var series = data.series || [];
    if (!labels.length || !series.length) {
      container.innerHTML = '<p class="chart-empty">لا توجد بيانات كافية لعرض الرسم</p>';
      return;
    }
    var root = makeRoot(container);
    var width = root.width;
    var height = root.height;
    var padTop = 24;
    var padBottom = 42;
    var padSide = 58;
    var plotW = width - padSide * 2;
    var plotH = height - padTop - padBottom;

    var max = 0;
    series.forEach(function (s) {
      (s.values || []).forEach(function (v) { max = Math.max(max, Number(v) || 0); });
    });
    max = niceMax(max);

    var gridCount = 4;
    for (var g = 0; g <= gridCount; g++) {
      var y = padTop + (plotH / gridCount) * g;
      root.svg.appendChild(svgEl("line", {
        x1: padSide, y1: y, x2: width - padSide, y2: y, stroke: "#eef2f7", "stroke-width": 1,
      }));
      var label = svgEl("text", { x: padSide - 10, y: y + 4, "text-anchor": "end", class: "chart-axis", "font-size": "11" });
      label.textContent = shortNum(max - (max / gridCount) * g);
      root.svg.appendChild(label);
    }

    var groupW = plotW / labels.length;
    var barW = Math.min(26, (groupW * 0.62) / series.length);
    labels.forEach(function (name, i) {
      var center = padSide + groupW * i + groupW / 2;
      series.forEach(function (item, si) {
        var value = Number((item.values || [])[i]) || 0;
        var barH = (plotH * value) / max;
        var x = center - (barW * series.length) / 2 + si * barW + 2;
        var rect = svgEl("rect", {
          x: x, y: padTop + plotH - barH, width: Math.max(barW - 4, 3), height: Math.max(barH, 1),
          rx: 5, fill: item.color, opacity: 0.92,
        });
        var title = svgEl("title");
        title.textContent = (item.name || "") + ": " + fmt(value);
        rect.appendChild(title);
        root.svg.appendChild(rect);
      });
      var xLabel = svgEl("text", { x: center, y: height - padBottom + 22, "text-anchor": "middle", class: "chart-axis", "font-size": "11" });
      xLabel.textContent = name.length > 10 ? name.slice(0, 10) + "…" : name;
      var xTitle = svgEl("title");
      xTitle.textContent = name;
      xLabel.appendChild(xTitle);
      root.svg.appendChild(xLabel);
    });
    drawLegend(container, series);
  }

  function donutChart(container, data) {
    var labels = data.labels || [];
    var values = data.values || [];
    var total = values.reduce(function (a, b) { return a + (Number(b) || 0); }, 0);
    var colors = data.colors || ["#0d9488", "#f59e0b", "#e11d48", "#1d4ed8", "#7c3aed"];
    if (!total) {
      container.innerHTML = '<p class="chart-empty">لا توجد بيانات</p>';
      return;
    }
    var root = makeRoot(container);
    var width = root.width;
    var height = root.height;
    var cx = width / 2;
    var cy = height / 2 - 6;
    var radius = Math.min(width, height) / 2 - 34;
    var inner = radius * 0.64;
    var angle = -Math.PI / 2;

    values.forEach(function (value, i) {
      var portion = (Number(value) || 0) / total;
      var end = angle + portion * Math.PI * 2;
      var large = portion > 0.5 ? 1 : 0;
      var x1 = cx + radius * Math.cos(angle);
      var y1 = cy + radius * Math.sin(angle);
      var x2 = cx + radius * Math.cos(end);
      var y2 = cy + radius * Math.sin(end);
      var ix2 = cx + inner * Math.cos(end);
      var iy2 = cy + inner * Math.sin(end);
      var ix1 = cx + inner * Math.cos(angle);
      var iy1 = cy + inner * Math.sin(angle);
      var path = "M" + x1 + "," + y1 +
        " A" + radius + "," + radius + " 0 " + large + " 1 " + x2 + "," + y2 +
        " L" + ix2 + "," + iy2 +
        " A" + inner + "," + inner + " 0 " + large + " 0 " + ix1 + "," + iy1 + " Z";
      var slice = svgEl("path", { d: path, fill: colors[i % colors.length], opacity: 0.94 });
      var title = svgEl("title");
      title.textContent = labels[i] + ": " + fmt(value);
      slice.appendChild(title);
      root.svg.appendChild(slice);
      angle = end;
    });

    var centerText = svgEl("text", { x: cx, y: cy + 4, "text-anchor": "middle", "font-size": "24", "font-weight": "800", fill: "#0f2a4a" });
    centerText.textContent = fmt(total);
    root.svg.appendChild(centerText);
    var centerSub = svgEl("text", { x: cx, y: cy + 26, "text-anchor": "middle", "font-size": "11", fill: "#94a3b8" });
    centerSub.textContent = "إجمالي الأصناف";
    root.svg.appendChild(centerSub);

    var legend = document.createElement("div");
    legend.className = "chart-legend chart-legend-col";
    labels.forEach(function (name, i) {
      var span = document.createElement("span");
      span.className = "chart-legend-item";
      var dot = document.createElement("i");
      dot.style.background = colors[i % colors.length];
      span.appendChild(dot);
      span.appendChild(document.createTextNode(name + " · " + fmt(values[i])));
      legend.appendChild(span);
    });
    container.appendChild(legend);
  }

  function renderAll() {
    document.querySelectorAll("[data-chart]").forEach(function (node) {
      var type = node.dataset.chart;
      var payload = node.querySelector("script[type='application/json']");
      if (!payload) return;
      var data;
      try {
        data = JSON.parse(payload.textContent);
      } catch (err) {
        return;
      }
      if (type === "line") lineChart(node, data);
      else if (type === "bar") barChart(node, data);
      else if (type === "donut") donutChart(node, data);
    });
  }

  var resizeTimer = null;
  window.addEventListener("resize", function () {
    clearTimeout(resizeTimer);
    resizeTimer = setTimeout(renderAll, 220);
  });

  if (document.readyState === "loading") {
    document.addEventListener("DOMContentLoaded", renderAll);
  } else {
    renderAll();
  }

  window.AlaslyCharts = { render: renderAll };
})();
