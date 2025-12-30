const FONT_KR = "Malgun Gothic";

exports.handler = async (event) => {
  if (event.httpMethod !== "POST") {
    return { statusCode: 405, body: "Method Not Allowed" };
  }

  let PptxGenJS;
  try {
    PptxGenJS = require("pptxgenjs");
  } catch (err) {
    console.error("pptxgenjs load fail:", err);
    return { statusCode: 500, body: "pptxgenjs 모듈을 불러오지 못했습니다." };
  }

  let payload = {};
  try {
    payload = JSON.parse(event.body || "{}");
  } catch {
    return { statusCode: 400, body: "Invalid JSON" };
  }

  const {
    notice_no = "",
    apt_name = "",
    period = "YYYY-MM-DD ~ YYYY-MM-DD",
    title = "제목을 입력하세요",
    body = ["(AI 본문 자리)"],
    footer = "",
    type = "general",
  } = payload;

  const [startRaw = "YYYY-MM-DD", endRaw = "YYYY-MM-DD"] = (period || "")
    .split("~")
    .map((s) => (s || "").trim());
  const start = startRaw.replace(/\./g, "-") || "YYYY-MM-DD";
  const end = endRaw.replace(/\./g, "-") || "YYYY-MM-DD";

  try {
    const pptx = new PptxGenJS();
    const mm = (v) => v / 25.4;
    const pageW = mm(210);
    const pageH = mm(297);
    pptx.defineLayout({ name: "A4", width: pageW, height: pageH });
    pptx.layout = "A4";

    const BLUE = "1e73c8";
    const DARK = "0c4a99";
    const TEXT = "1d1d1f";
    const LIGHT = "f5f8fd";
    const OUTLINE = "cdd9e7";

    const margin = mm(12);
    const gap = mm(4);
    const headerH = mm(18);
    const infoH = mm(12);
    const footerH = mm(16);
    const marginBottom = mm(10);
    const innerW = pageW - margin * 2;
    const infoTop = margin + headerH;
    const bodyTop = infoTop + infoH + gap;
    const bodyH = pageH - marginBottom - footerH - bodyTop;

    const slide = pptx.addSlide();

    const headerFontSize = 20;
    const footerFontSize = 18;

    // header bar
    slide.addShape(pptx.ShapeType.rect, {
      x: margin,
      y: margin,
      w: innerW,
      h: headerH,
      fill: BLUE,
      line: { color: BLUE, width: 0 },
    });
    slide.addText(title || "제목을 입력하세요", {
      x: margin,
      y: margin,
      w: innerW,
      h: headerH,
      fontSize: headerFontSize,
      fontFace: FONT_KR,
      color: "FFFFFF",
      bold: true,
      align: "center",
      valign: "middle",
      autoFit: true,
    });

    // info bar
    slide.addShape(pptx.ShapeType.rect, {
      x: margin,
      y: infoTop,
      w: innerW,
      h: infoH,
      fill: DARK,
      line: { color: DARK, width: 0 },
    });
    const infoTxt = `공고번호: ${notice_no || "-"}   |   게시기간: ${start} ~ ${end}`;
    slide.addText(infoTxt, {
      x: margin,
      y: infoTop,
      w: innerW,
      h: infoH,
      fontSize: 10,
      fontFace: FONT_KR,
      color: "FFFFFF",
      align: "center",
      valign: "middle",
      autoFit: true,
    });

    // body background
    slide.addShape(pptx.ShapeType.rect, {
      x: margin,
      y: bodyTop,
      w: innerW,
      h: bodyH,
      fill: LIGHT,
      line: { color: OUTLINE, width: 0.75 },
    });

  const bodyLinesRaw = Array.isArray(body)
    ? body.map((l) => (l ?? "").toString())
    : typeof body === "string"
    ? body.split(/\r?\n/)
    : [];
  const safeBodyLines =
    bodyLinesRaw.length > 0
      ? bodyLinesRaw
      : [
          "국토교통부 고시 [주택관리업자 및 사업자 선정지침] 제11조에 따라 사업자 선정결과를 아래와 같이 공고합니다.",
          "상호: {상호}",
          "주소: {주소}",
          "대표자: {대표자}",
          "연락처: {연락처}",
          "사업자등록번호: {사업자등록번호}",
          "계약금액: {계약금액}",
          "계약기간: {계약기간}",
          "계약사유: {계약사유}",
          "사유 및 결과: {사유및결과}",
        ];

    const baseFontSize = type === "general" ? 15 : 13;
    const lineSpace = Math.round(baseFontSize * 1.45);

    const bodyX = margin + mm(4);
    const bodyW = innerW - mm(8);
    const bodyY = bodyTop + mm(4);
    const bodyInnerH = bodyH - mm(8);

    if (type === "result") {
    const introLines = [];
    const tableLines = [];
    const footerLines = [];
    let centerLine = "";
    let tableStarted = false;

    safeBodyLines.forEach((line) => {
      if (!line.trim()) return;
      const trimmed = line.trim();
      if (trimmed === "- 아래 -") {
        centerLine = trimmed;
        tableStarted = true;
        return;
      }
      if (trimmed.includes(":")) {
        const [k, ...rest] = trimmed.split(":");
        const key = (k || "").trim();
        const val = rest.join(":").trim();
        if (key === "사유 및 결과") {
          footerLines.push(val || "(입력 없음)");
          tableStarted = true;
        } else {
          tableLines.push([key || "-", val || "-"]);
          tableStarted = true;
        }
      } else if (!tableStarted) {
        introLines.push(trimmed);
      } else {
        footerLines.push(trimmed);
      }
    });

      let cursorY = bodyY;
      if (introLines.length) {
        const introText = introLines.join("\n");
        const introHeight = Math.min(bodyInnerH * 0.3, Math.max(mm(18), mm(introLines.length * 8 + 8)));
        slide.addText(introText, {
          x: bodyX,
          y: cursorY,
          w: bodyW,
          h: introHeight,
          fontSize: 14,
          fontFace: FONT_KR,
          color: TEXT,
          align: "left",
          valign: "top",
          lineSpacing: lineSpace,
          paraSpaceAfter: 3,
          autoFit: true,
          wrap: true,
        });
        cursorY += introHeight + mm(3);
      }

      if (centerLine) {
        const centerHeight = mm(8);
        slide.addText(centerLine, {
          x: bodyX,
          y: cursorY,
          w: bodyW,
          h: centerHeight,
          fontSize: 14,
          fontFace: FONT_KR,
          color: TEXT,
          align: "center",
          valign: "middle",
          bold: true,
        });
        cursorY += centerHeight + mm(3);
      }

    let tableRows = tableLines.length ? tableLines : [];

      if (!tableRows.length) {
        tableRows = [
          ["상호", ""],
          ["주소", ""],
          ["대표자", ""],
          ["연락처", ""],
          ["사업자등록번호", ""],
          ["계약금액", ""],
          ["계약기간", ""],
          ["계약사유", ""],
        ];
      }

      const remainingH = bodyInnerH - (cursorY - bodyY);
      const tableHeight = Math.min(remainingH, mm(tableRows.length * 12 + 18));
      slide.addTable(
        [
          ["항목", "내용"],
          ...tableRows,
        ].map((r, idx) =>
          r.map((text, colIdx) => ({
            text,
            options: {
              fontFace: FONT_KR,
              fontSize: baseFontSize,
              bold: idx === 0 || colIdx === 0,
              color: TEXT,
              fill: idx === 0 ? "e8edf5" : "ffffff",
              border: { type: "solid", color: OUTLINE, pt: 1 },
              align: "center",
              valign: "middle",
            },
          }))
        ),
        {
          x: bodyX,
          y: cursorY,
          w: bodyW,
          h: tableHeight,
          colW: [mm(38), bodyW - mm(38)],
          margin: 2,
        }
      );
      cursorY += tableHeight + mm(4);

      if (footerLines.length && cursorY < bodyY + bodyInnerH - mm(10)) {
        const labelH = mm(6);
        const footHeight = Math.min(bodyY + bodyInnerH - cursorY - labelH, mm(Math.max(24, footerLines.length * 8 + 8)));
        if (footHeight > 0) {
          slide.addText("사유 및 결과", {
            x: bodyX,
            y: cursorY,
            w: bodyW,
            h: labelH,
            fontSize: 14,
            fontFace: FONT_KR,
            color: TEXT,
            align: "left",
            valign: "top",
            bold: true,
          });
          cursorY += labelH + mm(2);
          slide.addShape(pptx.ShapeType.rect, {
            x: bodyX,
            y: cursorY,
            w: bodyW,
            h: footHeight,
            fill: "f7f9fc",
            line: { color: OUTLINE, width: 1 },
          });
          slide.addText(footerLines.join("\n"), {
            x: bodyX + mm(2),
            y: cursorY + mm(2),
            w: bodyW - mm(4),
            h: footHeight - mm(4),
            fontSize: 14,
            fontFace: FONT_KR,
            color: TEXT,
            align: "left",
            valign: "top",
            lineSpacing: lineSpace,
            paraSpaceAfter: 3,
            autoFit: true,
            wrap: true,
          });
        }
      }
    } else {
      const bodyText = bodyLinesRaw.join("\n") || "(AI 본문 자리)";
      slide.addText(bodyText, {
        x: bodyX,
        y: bodyY,
        w: bodyW,
        h: bodyInnerH,
        fontSize: baseFontSize,
        fontFace: FONT_KR,
        color: TEXT,
        align: "left",
        valign: "top",
        lineSpacing: lineSpace,
        paraSpaceAfter: 3,
        autoFit: true,
        wrap: true,
      });
    }

    // footer
    const footerText = footer?.trim() || (apt_name ? `${apt_name} 관리사무소장 [직인생략]` : "");
    slide.addShape(pptx.ShapeType.rect, {
      x: margin,
      y: pageH - margin - footerH,
      w: innerW,
      h: footerH,
      fill: BLUE,
      line: { color: BLUE, width: 0 },
    });
    slide.addText(footerText, {
      x: margin,
      y: pageH - margin - footerH,
      w: innerW,
      h: footerH,
      fontSize: footerFontSize,
      fontFace: FONT_KR,
      color: "FFFFFF",
      bold: true,
      align: "center",
      valign: "middle",
      autoFit: true,
    });

    const pptxBuf = await pptx.write({ outputType: "nodebuffer" });
    const downloadName = type === "result" ? "결과안내문.pptx" : "notice_a4.pptx";

    return {
      statusCode: 200,
      headers: {
        "Content-Type": "application/vnd.openxmlformats-officedocument.presentationml.presentation",
        "Content-Disposition": `attachment; filename="${downloadName}"`,
      },
      body: pptxBuf.toString("base64"),
      isBase64Encoded: true,
    };
  } catch (err) {
    console.error("generate-pptx error:", err);
    return {
      statusCode: 500,
      body: `generate-pptx 실패: ${err && err.message ? err.message : err}`,
    };
  }
};
