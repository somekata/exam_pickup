/*********************************************************
 * ページTOPボタン制御
 *********************************************************/
document.addEventListener("DOMContentLoaded", () => {
  const btn = document.getElementById("backToTopBtn");
  if (!btn) return;

  // スクロール位置で表示/非表示
  window.addEventListener("scroll", () => {
    const y = window.scrollY || document.documentElement.scrollTop;
    btn.style.display = (y > 300) ? "block" : "none";
  });

  // クリックでスムーズスクロール
  btn.addEventListener("click", () => {
    window.scrollTo({ top: 0, behavior: "smooth" });
  });
});
