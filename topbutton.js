document.addEventListener('DOMContentLoaded', () => {
    // パイ型スクロール進捗インジケーターの要素を取得
    const progressPie = document.getElementById('scroll-progress-pie');

    // スクロール量の閾値（この量スクロールしたらインジケーターを表示）
    const scrollThreshold = 50; 

    // --- 1. スクロール進捗の計算と表示 ---
    const updateProgressPie = () => {
        const currentScroll = window.scrollY;
        const maxScroll = document.documentElement.scrollHeight - window.innerHeight;

        let progress = 0; // 進捗率 (0から1)
        if (maxScroll > 0) {
            progress = currentScroll / maxScroll;
        }

        // conic-gradient を使ってパイ型に塗りつぶす
        // #007bff: 進捗色, #eee: ベース色
        progressPie.style.backgroundImage = `conic-gradient(#007bff ${progress * 360}deg, #eee ${progress * 360}deg)`;

        // 最下部に達した時の特別なスタイル（例: 赤色に変化）
        if (progress >= 0.99) { // ほぼ最下部と判断
            progressPie.classList.add('fully-scrolled');
        } else {
            progressPie.classList.remove('fully-scrolled');
        }

        // スクロール量に基づいてインジケーターの表示・非表示を切り替え
        if (currentScroll > scrollThreshold) {
            progressPie.style.display = 'block';
            progressPie.style.opacity = '1';
            progressPie.style.cursor = 'pointer'; // マウスオーバー時にクリック可能であることを示す
        } else {
            progressPie.style.opacity = '0';
            progressPie.style.cursor = 'default';
            // フェードアウト後に非表示にする
            setTimeout(() => {
                if (parseFloat(progressPie.style.opacity) === 0) {
                    progressPie.style.display = 'none';
                }
            }, 300);
        }
    };

    // --- 2. クリックイベントの処理：TOPへスクロール ---
    progressPie.addEventListener('click', () => {
        // TOPへスムーズにスクロール
        window.scrollTo({
            top: 0, 
            behavior: 'smooth' 
        });
    });

    // イベントリスナーの設定と初期実行
    window.addEventListener('scroll', updateProgressPie);
    updateProgressPie(); // 初期状態を設定
});