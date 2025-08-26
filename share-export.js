// share-export.js
// ScoreAnalyzer.exportAsHtml() 오버라이드: ZIP 없이 "새 창" + 정적 스냅샷 즉시 표시 + 원본처럼 동작
(function () {
  if (!window.ScoreAnalyzer) return;

  ScoreAnalyzer.prototype.exportAsHtml = async function () {
    try {
      if (!this.combinedData) {
        alert('먼저 "분석 시작"으로 분석을 완료해 주세요.');
        return;
      }

      // 1) 정적 스냅샷: 현재 #results 복제 + 차트 캔버스→이미지
      const results = document.getElementById('results');
      const clone = results.cloneNode(true);
      clone.style.display = '';
      const origCanvases = results.querySelectorAll('canvas');
      const cloneCanvases = clone.querySelectorAll('canvas');
      cloneCanvases.forEach((c, i) => {
        const src = origCanvases[i];
        if (!src) return;
        try {
          const url = src.toDataURL('image/png');
          const img = new Image();
          img.src = url;
          img.width = src.width;
          img.height = src.height;
          c.replaceWith(img);
        } catch (_) {}
      });
      const snapshotHTML = clone.outerHTML;

      // 2) 현재 script.js를 인라인으로 삽입(캐시/경로 이슈 제거)
      let currentScriptText = '';
      try {
        const resp = await fetch('script.js', { cache: 'no-store' });
        currentScriptText = await resp.text();
      } catch (e) {
        console.warn('script.js 불러오기 실패:', e);
        currentScriptText = '';
      }

      // 3) 외부 라이브러리 + 스타일
      const XLSX_URL       = 'https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js';
      const CHART_URL      = 'https://cdn.jsdelivr.net/npm/chart.js';
      const DATALABELS_URL = 'https://cdn.jsdelivr.net/npm/chartjs-plugin-datalabels@2';
      const JSZIP_URL      = 'https://cdnjs.cloudflare.com/ajax/libs/jszip/3.10.1/jszip.min.js';
      const JSPDF_URL      = 'https://cdnjs.cloudflare.com/ajax/libs/jspdf/2.5.1/jspdf.umd.min.js';
      const H2C_URL        = 'https://cdnjs.cloudflare.com/ajax/libs/html2canvas/1.4.1/html2canvas.min.js';
      const CSS_URL_ABS    = 'https://ironmins.github.io/school_transcript_analysis/style.css'; // 배포/미리보기 모두 안전
      const CSS_URL_REL    = '../style.css'; // /reports/에 저장한 뒤에도 동작하도록 추가

      // 4) 공유용 HTML 조립 (+ SHARE_MODE 플래그 설정, init 이벤트리스너 무력화)
      const preloaded = JSON.stringify(this.combinedData);
      const html = `<!DOCTYPE html>
<html lang="ko">
<head>
  <meta charset="UTF-8"/>
  <meta name="viewport" content="width=device-width,initial-scale=1.0"/>
  <title>(공유) 고1 내신 분석 결과</title>
  <link rel="stylesheet" href="${CSS_URL_ABS}">
  <link rel="stylesheet" href="${CSS_URL_REL}">
  <script src="${XLSX_URL}"></script>
  <script src="${CHART_URL}"></script>
  <script src="${DATALABELS_URL}"></script>
  <script src="${JSZIP_URL}"></script>
  <script src="${JSPDF_URL}"></script>
  <script src="${H2C_URL}"></script>
  <style>
    body{max-width:1200px;margin:24px auto;padding:0 12px}
    .upload-section,#loading,#error{display:none !important} /* 공유에서는 업로드 UI 숨김 */
    .share-toolbar{display:flex;gap:8px;justify-content:flex-end;margin:8px 0 12px}
    .share-toolbar button{padding:8px 12px;border:1px solid #e5e7eb;border-radius:8px;background:#fff;cursor:pointer}
    .share-toolbar button:hover{background:#f3f4f6}
  </style>
</head>
<body>
  <!-- 즉시 사용 가능한 도구 -->
  <div class="share-toolbar">
    <button id="btnPrint">인쇄</button>
    <button id="btnPdf">PDF 저장</button>
  </div>

  <!-- 1차: 정적 스냅샷을 즉시 표시(데이터 보장) -->
  ${snapshotHTML}

  <!-- 공유 모드 플래그 -->
  <script>window.__SHARE_MODE__ = true; window.PRELOADED_DATA = ${preloaded};</script>

  <!-- 2차: 데이터 주입 + 현재 script.js 인라인 삽입 후 원본처럼 재렌더 -->
  ${currentScriptText ? `<script>\n${currentScriptText.replace(/<\/script>/g,'<\\/script>')}\n<\/script>` : ''}

  <!-- 3차: 강제 초기화(DOM 상태와 무관하게) + 탭 보정 + PDF/인쇄 + 이벤트리스너 no-op -->
  <script>
    (function start(){
      function boot(){
        try{
          // 공유 모드에서는 업로드/드래그 등 초기 이벤트리스너가 필요 없음 → 안전하게 무력화
          if (window.__SHARE_MODE__ && window.ScoreAnalyzer) {
            const orig = ScoreAnalyzer.prototype.initializeEventListeners;
            ScoreAnalyzer.prototype.initializeEventListeners = function(){ /* no-op in share view */ };
          }
          // 업로드 UI 숨기고 결과 표시
          const up = document.querySelector('.upload-section'); if (up) up.style.display = 'none';
          const rs = document.getElementById('results'); if (rs) rs.style.display = 'block';

          if (window.ScoreAnalyzer) {
            window.scoreAnalyzer = new ScoreAnalyzer();
            // 숨은 탭 차트 크기 보정(왕복 클릭)
            setTimeout(() => {
              const clickTab = (name) => document.querySelector('.tab-btn[data-tab="'+name+'"]')?.click();
              clickTab('grade-analysis');
              setTimeout(() => clickTab('subjects'), 60);
            }, 120);
          }
        }catch(e){ console.error(e); }
      }
      if (window.ScoreAnalyzer) boot();
      else window.addEventListener('load', boot, {once:true});

      // 인쇄
      document.getElementById('btnPrint')?.addEventListener('click', ()=>window.print());

      // PDF 저장(#results 캡처)
      document.getElementById('btnPdf')?.addEventListener('click', async ()=>{
        const el = document.getElementById('results'); if(!el) return;
        const { jsPDF } = window.jspdf;
        const canvas = await html2canvas(el, {scale:2, useCORS:true});
        const pdf = new jsPDF('p','pt','a4');
        const pageW = pdf.internal.pageSize.getWidth();
        const pageH = pdf.internal.pageSize.getHeight();
        const imgW = pageW - 40;
        const ratio = imgW / canvas.width;
        const chunkH = (pageH - 40) / ratio;
        let srcY = 0;
        while (srcY < canvas.height) {
          const h = Math.min(chunkH, canvas.height - srcY);
          const pageCanvas = document.createElement('canvas');
          pageCanvas.width = canvas.width;
          pageCanvas.height = h;
          pageCanvas.getContext('2d').drawImage(canvas, 0, srcY, canvas.width, h, 0, 0, canvas.width, h);
          const pageImg = pageCanvas.toDataURL('image/png');
          if (srcY > 0) pdf.addPage();
          pdf.addImage(pageImg, 'PNG', 20, 20, imgW, h * ratio);
          srcY += h;
        }
        const stamp = new Date().toISOString().replace(/[:.]/g,'-');
        pdf.save('내신분석_공유_'+stamp+'.pdf');
      });
    })();
  </script>
</body>
</html>`;

      // 5) 새 창으로 열기
      const w = window.open('', '_blank');
      w.document.open(); w.document.write(html); w.document.close();

    } catch (e) {
      console.error(e);
      alert('공유용 새 창 생성 중 오류가 발생했습니다.');
    }
  };
})();