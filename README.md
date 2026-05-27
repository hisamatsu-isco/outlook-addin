# Attachment Warning v4

Outlook向け Smart Alerts アドイン。メール送信時にローカルファイル添付を検知し、OneDrive等のリンク共有を促す警告ポップアップを表示します。

**運用主体:** ISCO 情報システム
**対応環境:** Outlook on Mac / Windows / Web
**配信:** GitHub Pages + Microsoft 365 統合アプリ展開

---

## 機能概要

| 条件 | 動作 |
|------|------|
| ローカルファイル添付あり (PDF, Word, Excel等) | ⚠️ 警告ポップアップを表示 |
| OneDrive / SharePoint / box のリンク添付のみ | ✅ 警告なしで送信 |
| 添付なし | ✅ 警告なしで送信 |
| 添付取得失敗 | ✅ 警告なしで送信 (安全側に倒す) |

### 警告メッセージ

> ISCOルールとして、ファイルの直接添付は原則禁止していますので、OneDrive等へアップロードしての送信をお願いします。やむを得ない場合に限り、「そのまま送信する」としてください。

ボタン: `そのまま送信する` / `送信しない`

---

## 構成ファイル

```
attachment-warning/
├── manifest.xml          # アドインのマニフェスト (管理センターにアップロード)
├── launchevent.js        # Smart Alerts のイベントハンドラ (本体ロジック)
├── commands.html         # JavaScript-only ランタイムのエントリ HTML
├── taskpane.html         # (タスクペイン用、現状はサポートURL用)
├── taskpane.js           # 同上
└── assets/
    ├── icon-64.png
    └── icon-128.png
```

ホスティング先 (GitHub Pages):
```
https://hisamatsu-isco.github.io/outlook-addin/
```

---

## デプロイ方法 (manifest.xml 更新時)

1. ローカルで `manifest.xml` または `launchevent.js` を編集
2. `manifest.xml` の `<Version>` をインクリメント (例: `3.0.0.1` → `3.0.0.2`)
3. `git add -A && git commit && git push origin main`
4. GitHub Pages 反映を 1〜2 分待つ
5. [Microsoft 365 管理センター](https://admin.microsoft.com) → 設定 → 統合アプリ
6. **Attachment Warning v4** → **アドインの更新** → 新しい `manifest.xml` をアップロード
7. テナント反映を 5〜10 分待つ
8. シークレットウィンドウで Outlook Web を開き、添付付きメールで動作確認

### バージョン番号のルール

- マニフェスト構造/参照URL等の変更 → メジャー or マイナーをアップ
- メッセージ文言だけの修正 → パッチをアップ
- バージョンを上げ忘れると Microsoft 365 側がアドインを更新と認識しないので注意

---

## 開発履歴

### v3.0.0.0 〜 v3.0.0.2 (2026年5月)

最初の組織展開で立ち上げ。下記のハマりポイントを順に解消して安定動作に到達。

| バージョン | 変更点 |
|----------|--------|
| 3.0.0.0  | 初期展開、Runtimesブロック追加 |
| 3.0.0.1  | `Office.onReady()` のコールバック方式修正、不要な commandId 削除 |
| 3.0.0.2  | `attachments.getAsync()` → `getAttachmentsAsync()` に修正 (真犯人) |
| 3.0.0.2  | デバッグログ削除、警告メッセージを ISCO ルール文言に更新 |

---

## ハマりポイント (将来のメンテ用)

これは超重要です。同じ罠にハマらないように残します。

### 1. `Office.onReady()` の呼び方

❌ ダメな書き方:
```javascript
Office.onReady();
Office.actions.associate("onMessageSendHandler", handler);
```

✅ 正しい書き方:
```javascript
Office.onReady(function() {
  Office.actions.associate("onMessageSendHandler", handler);
});
```

`Office.onReady()` をコールバックなしで呼ぶと、Office.js の初期化完了を待たずに `associate` が走り、ハンドラが正しく登録されない。症状としてはポップアップは出るが「メッセージを処理しています」のままで `event.completed()` が呼ばれない。

### 2. 添付ファイル取得APIの違い

❌ Smart Alerts の launchevent コンテキストでは動かない:
```javascript
Office.context.mailbox.item.attachments.getAsync(callback);
```
→ `Cannot read properties of undefined (reading 'getAsync')` エラー。`attachments` プロパティがそもそも未定義になっている。

✅ Smart Alerts では必ずこちら:
```javascript
Office.context.mailbox.item.getAttachmentsAsync(callback);
```

Office.js のドキュメントでは両方の書き方が混在しているが、**LaunchEvent (Smart Alerts) ランタイムでは `getAttachmentsAsync()` メソッドを使う**。

### 3. manifest.xml の Runtimes ブロック

OnMessageSend を使う場合、`<Runtimes>` で JavaScript-only ランタイムを定義する必要がある。

```xml
<Runtimes>
  <Runtime resid="WebViewRuntime.Url">
    <Override type="javascript" resid="JSRuntime.Url"/>
  </Runtime>
</Runtimes>
```

そして `WebViewRuntime.Url` は `commands.html` を、`JSRuntime.Url` は `launchevent.js` を指す。`commands.html` の中身も忘れずに用意すること:

```html
<!DOCTYPE html>
<html><head>
  <script src="https://appsforoffice.microsoft.com/lib/1/hosted/office.js"></script>
  <script src="launchevent.js"></script>
</head><body></body></html>
```

### 4. Office.js のスクリプトキャッシュ

`launchevent.js` の中身を変えてプッシュしても、Outlook 側で**24時間ほど古いコードがキャッシュされ続ける**ことがある。ブラウザのキャッシュクリアやシークレットモードでも消えない (Microsoft 側のキャッシュなので)。

#### 確実にキャッシュを破棄する方法
1. `manifest.xml` の `<Version>` を上げる
2. 管理センターで「アドインの更新」を実行

それでもダメなら:
3. `launchevent.js` をファイル名ごとリネーム (例: `launchevent-v5.js`)
4. `commands.html` の参照を更新
5. `manifest.xml` の `<bt:Url id="JSRuntime.Url" ...>` も更新
6. 再度バージョンを上げて管理センターで更新

### 5. manifest.xml の検証

公式ツールで構造チェックができる。怪しいと思ったら走らせる:

```bash
npx office-addin-manifest validate manifest.xml
```

### 6. デバッグ方法

`launchevent.js` に `console.log` を仕込んで、Outlook Web で F12 → Console → フィルタに識別子を入れて確認できる。シークレットウィンドウだとキャッシュの影響を受けにくい。

```javascript
console.log("[AttachmentWarning] launchevent.js loaded");
Office.onReady(function() {
  console.log("[AttachmentWarning] Office.onReady fired");
  // ...
});
```

ログが出る場所まで確認することで、どこで止まっているか特定できる。

### 7. ファイル選択ダイアログで manifest.xml がグレーアウト

管理センターの「アドインの更新」で `manifest.xml` を選ぼうとしてグレーアウトする場合、ダイアログ左下の「**オプションを表示**」をクリックしてファイル形式フィルタを変更するか、Macなら `Cmd + Shift + .` で隠しファイル表示を切り替えると選べることがある。

---

## ユーザー割り当ての変更

組織展開のスコープ変更は管理センターから:

1. [Microsoft 365 管理センター](https://admin.microsoft.com) → 設定 → 統合アプリ
2. Attachment Warning v4 → **ユーザー** タブ
3. ラジオボタン: `自分だけ` / `組織全体` / `特定のユーザーまたはグループ`
4. 反映には数時間〜最大24時間かかることがある

---

## ライセンス・運用

- 内部利用のみ
- アドイン ID: `ad4b4bbc-83de-48e2-8fb4-6f8e4bb8a643`
- メンテナンス: ISCO 情報システム
- 開発記録: 2026年5月25〜26日(2日間の死闘の末完成)

---

## 関連リソース

- [Office.js LaunchEvent API リファレンス](https://learn.microsoft.com/en-us/javascript/api/outlook/office.launchevent)
- [Outlook Smart Alerts 公式ドキュメント](https://learn.microsoft.com/en-us/office/dev/add-ins/outlook/smart-alerts-onmessagesend-walkthrough)
- [Office Add-in Manifest 検証ツール](https://www.npmjs.com/package/office-addin-manifest)
