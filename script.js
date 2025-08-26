
// 공유 분석 페이지 생성 함수
function openShareableAnalysisPage(scoreAnalyzer) {
    const data = scoreAnalyzer.data;

    const newWindow = window.open("", "_blank");
    if (!newWindow) return;

    const htmlContent = `
<!DOCTYPE html>
<html lang="ko">
<head>
    <meta charset="UTF-8">
    <title>공유용 성적 분석 페이지</title>
    <style>
        body { font-family: Arial, sans-serif; margin: 20px; }
        .tab { cursor: pointer; padding: 10px 20px; margin-right: 10px; background: #eee; display: inline-block; }
        .tab.active { background: #ccc; }
        .tab-content { display: none; margin-top: 20px; }
        .tab-content.active { display: block; }
        table { border-collapse: collapse; width: 100%; margin-top: 10px; }
        th, td { border: 1px solid #999; padding: 8px; text-align: center; }
    </style>
</head>
<body>
    <h1>공유용 성적 분석</h1>
    <div>
        <div class="tab active" data-tab="subject">과목별 분석</div>
        <div class="tab" data-tab="distribution">평균등급 분포</div>
        <div class="tab" data-tab="student">학생별 분석</div>
    </div>

    <div id="subject" class="tab-content active">
        <h2>과목별 분석</h2>
        <table>
            <thead><tr><th>과목명</th><th>평균</th><th>표준편차</th></tr></thead>
            <tbody>
                ${data.subjectStats.map(s => `<tr><td>${s.subject}</td><td>${s.average}</td><td>${s.stddev}</td></tr>`).join("")}
            </tbody>
        </table>
    </div>

    <div id="distribution" class="tab-content">
        <h2>평균등급 분포</h2>
        <table>
            <thead><tr><th>등급</th><th>학생 수</th></tr></thead>
            <tbody>
                ${Object.entries(data.gradeDistribution).map(([grade, count]) => `<tr><td>${grade}</td><td>${count}</td></tr>`).join("")}
            </tbody>
        </table>
    </div>

    <div id="student" class="tab-content">
        <h2>학생별 분석</h2>
        <table>
            <thead><tr><th>이름</th><th>총점</th><th>평균</th></tr></thead>
            <tbody>
                ${data.students.map(s => `<tr><td>${s.name}</td><td>${s.total}</td><td>${s.average}</td></tr>`).join("")}
            </tbody>
        </table>
    </div>

    <script>
        const tabs = document.querySelectorAll(".tab");
        const contents = document.querySelectorAll(".tab-content");

        tabs.forEach(tab => {
            tab.addEventListener("click", () => {
                tabs.forEach(t => t.classList.remove("active"));
                contents.forEach(c => c.classList.remove("active"));
                tab.classList.add("active");
                document.getElementById(tab.dataset.tab).classList.add("active");
            });
        });
    </script>
</body>
</html>
    `;

    newWindow.document.write(htmlContent);
    newWindow.document.close();
}
