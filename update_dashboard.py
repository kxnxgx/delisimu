import pandas as pd
import json
import os
import sys

# Windows環境での文字化け対策
sys.stdout.reconfigure(encoding='utf-8')

def generate_dashboard():
    print("Kankenダッシュボード生成を開始します...")
    
    # === 【設定エリア】 ===========================
    SALES_RATIO = 158
    EXCEL_FILE = "Kanken23510_2026年版_v6.xlsx"
    OUTPUT_FILE = "Kanken_Dashboard_最新版.html"
    THRESHOLD_LARGE = -500 
    # ==============================================

    try:
        xl = pd.ExcelFile(EXCEL_FILE)
        df = xl.parse("全体サマリー", header=3)
        df.columns = df.columns.astype(str).str.replace('\n', '', regex=False).str.strip()
        df = df.dropna(subset=['カラー'])
        df = df[df['カラー'] != '合計'] 
    except Exception as e:
        print(f"【エラー】Excelファイルの読み込み中にエラーが発生しました。\n詳細: {e}")
        return

    try:
        total_supply = int(df['総供給(6月〜)'].sum())
        total_demand = int(df['需要予測(6〜12月)'].sum())
        total_diff = int(df['過不足'].sum())
        
        priority = []
        control_large = []
        control_small = []
        
        for _, row in df.iterrows():
            name = str(row['カラー'])
            diff = float(row['過不足'])
            method = str(row.get('配分方式', ''))
            
            if diff >= 0:
                priority.append({"name": name, "diff": diff})
            elif diff <= THRESHOLD_LARGE:
                control_large.append({"name": name, "method": method})
            else:
                control_small.append({"name": name, "method": method})
                
        # 指定された需要過多の4色
        shortage_colors = ["Fog-Pink", "Pink", "Ox Red", "Sky Blue-Light Oak"]
        
        # 需要過多の4色のデータを抽出（指定順に並べ替え）
        df_shortage = df[df['カラー'].isin(shortage_colors)]
        df_shortage = df_shortage.set_index('カラー').reindex(shortage_colors).reset_index()
        
        # それ以外のカラーから、過不足がプラスで、プラス幅が大きい上位3色を抽出
        df_surplus = df[~df['カラー'].isin(shortage_colors)]
        df_surplus = df_surplus[df_surplus['過不足'] > 0].sort_values('過不足', ascending=False).head(3)
        
        # 4色と3色を合算してグラフ用データとする
        chart_df = pd.concat([df_shortage, df_surplus]).dropna(subset=['カラー'])
        
        chart_labels = chart_df['カラー'].tolist()
        chart_supply = chart_df['総供給(6月〜)'].fillna(0).astype(int).tolist()
        chart_demand = chart_df['需要予測(6〜12月)'].fillna(0).astype(int).tolist()
        
        injected_data = {
            "kpi": {
                "sales_ratio": SALES_RATIO,
                "total_supply": total_supply,
                "total_demand": total_demand,
                "total_diff": total_diff
            },
            "priority": priority,
            "control_large": control_large,
            "control_small": control_small,
            "chart_labels": chart_labels,
            "chart_supply": chart_supply,
            "chart_demand": chart_demand
        }
        
        html_template = """<!DOCTYPE html>
<html lang="ja">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>2026年下半期 Kanken 出荷計画ダッシュボード</title>
    <!-- 万が一ロード失敗・ハングした場合のフェイルセーフ -->
    <style>body { visibility: visible !important; }</style>
    
    <script src="https://cdn.tailwindcss.com"></script>
    <script src="https://cdn.jsdelivr.net/npm/chart.js"></script>
    <script defer src="https://cdn.jsdelivr.net/npm/alpinejs@3.x.x/dist/cdn.min.js"></script>
    <style>
        body { font-family: 'Helvetica Neue', Arial, 'Hiragino Kaku Gothic ProN', 'Hiragino Sans', Meiryo, sans-serif; }
        [x-cloak] { display: none !important; }
    </style>
</head>
<body class="bg-slate-50 text-slate-800 h-screen flex flex-col" x-data="dashboardApp">

    <header class="bg-slate-900 text-white p-4 shadow-md flex justify-between items-center">
        <div>
            <h1 class="text-xl font-bold tracking-wider">Kanken (23510) 出荷計画 2026</h1>
            <p class="text-sm text-slate-300 mt-1">店舗スタッフ向け 共有ダッシュボード</p>
        </div>
        <div class="flex space-x-2">
            <button @click="activeTab = 'summary'" :class="{'bg-blue-600': activeTab === 'summary', 'bg-slate-700 hover:bg-slate-600': activeTab !== 'summary'}" class="px-4 py-2 rounded-md text-sm font-medium transition">
                1. 全体サマリー
            </button>
            <button @click="activeTab = 'action'" :class="{'bg-blue-600': activeTab === 'action', 'bg-slate-700 hover:bg-slate-600': activeTab !== 'action'}" class="px-4 py-2 rounded-md text-sm font-medium transition">
                2. カラー別接客方針
            </button>
        </div>
    </header>

    <main class="flex-1 p-6 overflow-auto">

        <!-- 共通KPIカード -->
        <div class="grid grid-cols-1 md:grid-cols-4 gap-4 mb-6">
            <div class="bg-white p-4 rounded-lg shadow-sm border border-slate-200 border-l-4 border-l-amber-500">
                <p class="text-sm text-slate-500 font-medium">1〜6月 販売数昨対比</p>
                <div class="flex items-baseline mt-1">
                    <p class="text-3xl font-bold text-slate-800" x-text="kpi.sales_ratio"></p><span class="text-lg font-bold text-slate-800">%</span>
                    <p class="ml-2 text-xs text-amber-700 bg-amber-100 px-2 py-0.5 rounded-full">需要過熱</p>
                </div>
            </div>
            <div class="bg-white p-4 rounded-lg shadow-sm border border-slate-200">
                <p class="text-sm text-slate-500 font-medium">総供給見込み (7月〜)</p>
                <p class="text-2xl font-bold text-slate-700 mt-1"><span x-text="kpi.total_supply.toLocaleString()"></span> <span class="text-sm font-normal text-slate-500">個</span></p>
            </div>
            <div class="bg-white p-4 rounded-lg shadow-sm border border-slate-200">
                <p class="text-sm text-slate-500 font-medium">需要予測 (7〜12月)</p>
                <p class="text-2xl font-bold text-slate-700 mt-1"><span x-text="kpi.total_demand.toLocaleString()"></span> <span class="text-sm font-normal text-slate-500">個</span></p>
            </div>
            <div :class="kpi.total_diff < 0 ? 'bg-red-50 border-l-red-500' : 'bg-emerald-50 border-l-emerald-500'" class="bg-white p-4 rounded-lg shadow-sm border border-slate-200 border-l-4">
                <p class="text-sm font-medium" :class="kpi.total_diff < 0 ? 'text-red-600' : 'text-emerald-600'">全体過不足</p>
                <div class="flex items-baseline mt-1">
                    <p class="text-3xl font-bold" :class="kpi.total_diff < 0 ? 'text-red-600' : 'text-emerald-600'">
                        <span x-text="kpi.total_diff > 0 ? '+' : ''"></span><span x-text="kpi.total_diff.toLocaleString()"></span> <span class="text-sm font-normal">個</span>
                    </p>
                    <p x-show="kpi.total_diff < 0" class="ml-2 text-xs text-red-700 bg-red-200 px-2 py-0.5 rounded-full">品薄警戒</p>
                    <p x-show="kpi.total_diff >= 0" class="ml-2 text-xs text-emerald-700 bg-emerald-200 px-2 py-0.5 rounded-full">在庫安定</p>
                </div>
            </div>
        </div>

        <!-- タブ1: 全体サマリー -->
        <div x-show="activeTab === 'summary'" class="bg-white p-6 rounded-lg shadow-sm border border-slate-200">
            <h2 class="text-xl font-bold border-b pb-2 mb-4 text-slate-800">「無い色」を嘆くのではなく、「ある色」でどう売上を作るかが勝負です。</h2>
            <div class="bg-blue-50 border-l-4 border-blue-500 p-4 mb-6 rounded">
                <p class="text-sm text-blue-800 font-medium">💡 需要予測データについて</p>
                <p class="text-sm text-blue-700 mt-1">
                    直近のトレンドを反映させるため**「2025年の総売上実績」**をベースにしつつ、欠品による波の乱れを補正するため**「正常だった2024年の月次構成比（売れる波）」**を掛け合わせて算出した、現場の実情に最も近いリアルな需要予測です。
                </p>
            </div>
            <p class="text-slate-600 mb-6">
                1〜6月の販売数は前年比<span x-text="kpi.sales_ratio" class="font-bold"></span>%と非常に高いお客様の需要を示していますが、皆様もご存知の通り、一部カラーで在庫切れが発生しており、供給が追いついていません。<br>
                下半期はこの「圧倒的な品薄」と向き合い、目当ての色が無いお客様をいかに<b>「在庫に余裕のあるカラー（代替色）へご案内できるか」</b>が、店舗の売上を左右する最大の鍵になります。
            </p>
            <div class="w-full" style="height: 400px;">
                <canvas id="summaryChart"></canvas>
            </div>
        </div>

        <!-- タブ2: カラー別接客方針 -->
        <div x-show="activeTab === 'action'" x-cloak>
            <div class="mb-4">
                <h2 class="text-xl font-bold text-slate-800">下半期の接客方針：「優先提案カラー」でお客様を逃さない！</h2>
                <p class="text-slate-600 mt-1">現在、本部側で全体の在庫状況に応じた「出荷コントロール（制限）」をかけています。現場では以下の「優先提案カラー」を中心に接客を展開してください。</p>
            </div>

            <!-- 【主役】優先提案カラー -->
            <div class="bg-white rounded-lg shadow-sm border-2 border-emerald-500 overflow-hidden mb-8">
                <div class="bg-emerald-500 text-white p-3 text-center">
                    <h3 class="font-bold text-lg tracking-wide">🎯 優先提案カラー (在庫潤沢・出荷制限なし)</h3>
                </div>
                <div class="p-5">
                    <p class="text-slate-700 mb-4 font-medium">需要に対して在庫が十分に確保できています。店頭のメインディスプレイに配置し、目当てのカラー（欠品色）がないお客様には、まずこの中から代替提案を行い絶対に機会損失を防いでください。</p>
                    
                    <div class="grid grid-cols-1 sm:grid-cols-2 md:grid-cols-4 gap-4">
                        <template x-for="color in priority" :key="color.name">
                            <div class="bg-emerald-50 border border-emerald-200 rounded p-3 flex flex-col justify-center items-center">
                                <span class="font-bold text-lg text-slate-800" x-text="color.name"></span>
                                <span class="text-emerald-600 font-bold text-xs mt-1" x-text="color.diff >= 300 ? '十分な余裕あり' : '余裕あり'"></span>
                            </div>
                        </template>
                    </div>
                </div>
            </div>

            <!-- 【脇役】本部のステータス共有 -->
            <h3 class="font-bold text-slate-600 mb-3 border-l-4 border-slate-400 pl-2">【参考】本部の出荷コントロール状況（入荷が制限されるカラー）</h3>
            <div class="grid grid-cols-1 md:grid-cols-2 gap-6">
                
                <!-- コントロール大 -->
                <div class="bg-slate-50 rounded-lg border border-slate-200 overflow-hidden">
                    <div class="bg-slate-200 text-slate-700 p-2 text-center border-b border-slate-300">
                        <h4 class="font-bold text-sm">⚠️ 出荷コントロール：大 (大幅不足)</h4>
                    </div>
                    <div class="p-4">
                        <p class="text-xs text-slate-500 mb-3">圧倒的に数が足りないため、本部の配分制限が強くかかります。入荷してもすぐ売り切れる前提で接客してください。</p>
                        <ul class="space-y-2 text-sm">
                            <template x-for="color in control_large" :key="color.name">
                                <li class="flex justify-between items-center border-b border-slate-200 pb-1">
                                    <span class="font-medium text-slate-700" x-text="color.name"></span>
                                    <span class="text-slate-400 text-xs" x-text="color.method"></span>
                                </li>
                            </template>
                        </ul>
                    </div>
                </div>

                <!-- コントロール小 -->
                <div class="bg-slate-50 rounded-lg border border-slate-200 overflow-hidden">
                    <div class="bg-slate-200 text-slate-700 p-2 text-center border-b border-slate-300">
                        <h4 class="font-bold text-sm">📊 出荷コントロール：小 (やや不足)</h4>
                    </div>
                    <div class="p-4">
                        <p class="text-xs text-slate-500 mb-3">やや数が足りない状況です。入荷タイミングでの需要に応じつつ、年間を通して展開できるよう制限をかけて配分します。</p>
                        <ul class="space-y-2 text-sm">
                            <template x-for="color in control_small" :key="color.name">
                                <li class="flex justify-between items-center border-b border-slate-200 pb-1">
                                    <span class="font-medium text-slate-700" x-text="color.name"></span>
                                    <span class="text-slate-400 text-xs" x-text="color.method"></span>
                                </li>
                            </template>
                        </ul>
                    </div>
                </div>

            </div>
        </div>
    </main>

    <script>
        // === 自動注入されたデータ ===
        const INJECTED_DATA = __JSON_DATA_HERE__;
        // ========================

        // Alpine.js のデータ定義
        document.addEventListener('alpine:init', () => {
            Alpine.data('dashboardApp', () => ({
                activeTab: 'summary',
                ...INJECTED_DATA
            }))
        })

        // グラフの描画 (Alpineから分離して確実に実行)
        document.addEventListener("DOMContentLoaded", function() {
            try {
                const ctx = document.getElementById('summaryChart');
                if (ctx && INJECTED_DATA.chart_labels) {
                    new Chart(ctx, {
                        type: 'bar',
                        data: {
                            labels: INJECTED_DATA.chart_labels,
                            datasets: [
                                {
                                    label: '下半期 総供給見込み',
                                    data: INJECTED_DATA.chart_supply,
                                    backgroundColor: '#94a3b8', 
                                },
                                {
                                    label: '下半期 需要予測',
                                    data: INJECTED_DATA.chart_demand,
                                    backgroundColor: '#f43f5e', 
                                }
                            ]
                        },
                        options: {
                            responsive: true,
                            maintainAspectRatio: false,
                            plugins: {
                                legend: { position: 'bottom' },
                                title: {
                                    display: true,
                                    text: '主要カラー別：下半期の供給見込み vs リアルな需要予測',
                                    font: { size: 16 }
                                }
                            },
                            scales: {
                                y: { beginAtZero: true, title: { display: true, text: '数量（個）' } }
                            }
                        }
                    });
                }
            } catch(e) {
                console.error("Chart error:", e);
            }
        });
    </script>
</body>
</html>"""

        # ensure_ascii=True で日本語をエスケープし、スクリプト破壊を防ぐ
        json_str = json.dumps(injected_data, ensure_ascii=True)
        final_html = html_template.replace("__JSON_DATA_HERE__", json_str)
        
        with open(OUTPUT_FILE, "w", encoding="utf-8") as f:
            f.write(final_html)
            
        print(f"ダッシュボードの自動生成が完了しました！\n出力ファイル: {OUTPUT_FILE}")
        
    except Exception as e:
        print(f"【エラー】データ処理中に予期せぬエラーが発生しました。\n詳細: {e}")

if __name__ == "__main__":
    generate_dashboard()
