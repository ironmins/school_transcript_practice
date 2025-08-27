function downloadResultPage() {
  if (!window.analysisHTML) {
    alert("먼저 분석을 완료해야 결과를 저장할 수 있습니다.");
    return;
  }

  const template = `
<!DOCTYPE html>
<html lang="ko">
<head>
  <meta charset="UTF-8">
  <title>분석 결과</title>
</head>
<body>
  <h1>고등학교 1학년 성적 분석 결과</h1>
  <div id="resultContainer">
    ${window.analysisHTML}
  </div>
</body>
</html>`;

  const blob = new Blob([template], { type: "text/html" });
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = "analysis_result.html";
  a.click();
  URL.revokeObjectURL(url);
}
