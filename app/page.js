"use client";

import { useEffect, useState } from "react";
import { PublicClientApplication } from "@azure/msal-browser";
import { ArrayContainer } from "./components/DataView";

const msalInstance = new PublicClientApplication({
  auth: {
    clientId: "d843a25b-962d-4912-9740-5d7f83ede221",
    authority: "https://login.microsoftonline.com/a7929d4b-ae1e-4e78-8ff8-3293161eac3a",
    redirectUri: "http://localhost:3000",
  },
});

const MOCK_SUMMARIES = [
  {
    id: "msg-001",
    subject: "【重要】第3四半期 売上レポートについて",
    from: "tanaka.hiroshi@example.co.jp",
    receivedAt: "2026-06-18T09:23:00Z",
    summary: "第3四半期の売上が前年比12%増となった旨の報告。添付のExcelファイルに詳細データあり。返信期限は6月20日。",
    tags: ["重要", "売上", "レポート"],
    attachments: [
      { name: "Q3_sales_report.xlsx", sizeKB: 248 },
      { name: "summary_slide.pptx", sizeKB: 1024 },
    ],
  },
  {
    id: "msg-002",
    subject: "来週の定例MTGについて",
    from: "sato.yuki@example.co.jp",
    receivedAt: "2026-06-18T11:05:00Z",
    summary: "来週月曜13:00からの定例MTGをZoomに変更する旨の連絡。URLは本文中に記載。",
    tags: ["MTG", "連絡"],
    attachments: [],
  },
  {
    id: "msg-003",
    subject: "新しいセキュリティポリシーの適用について",
    from: "it-support@example.co.jp",
    receivedAt: "2026-06-17T14:30:00Z",
    summary: "7月1日より全端末に新しいセキュリティポリシーが適用される。VPN接続の手順変更あり、添付マニュアルを確認すること。",
    tags: ["セキュリティ", "IT", "重要"],
    attachments: [
      { name: "vpn_manual_v2.pdf", sizeKB: 512 },
    ],
  },
];

export default function Page() {
  const [isLoggedIn, setIsLoggedIn] = useState(false);
  const [apiResult, setApiResult] = useState(MOCK_SUMMARIES);

  useEffect(() => {
    const init = async () => {
      await msalInstance.initialize();
      const res = await msalInstance.handleRedirectPromise();
      if (res) setIsLoggedIn(true);
      if (msalInstance.getAllAccounts().length > 0) setIsLoggedIn(true);
    };
    init();
  }, []);

  const login = async () => {
    await msalInstance.loginRedirect({
      scopes: ["openid", "profile", "email"],
    });
  };

  const callApi = async () => {
    const accounts = msalInstance.getAllAccounts();
    if (accounts.length === 0) { alert("ログインしていません"); return; }

    const tokenRes = await msalInstance.acquireTokenSilent({
      account: accounts[0],
      scopes: ["openid", "profile", "email"],
    });

    const res = await fetch(
      "https://pd3bdzc1wg.execute-api.ap-northeast-1.amazonaws.com/summaries",
      { headers: { Authorization: `Bearer ${tokenRes.idToken}` } }
    );
    const data = await res.json();
    setApiResult(data);
  };

  return (
    <>
      <header className="app-header">
        <div className="app-logo">
          <div className="app-logo-mark">✉</div>
          Mail Summary
        </div>
        <div className="app-header-right">
          {isLoggedIn ? (
            <>
              <div className="login-badge">
                <span className="login-badge-dot" />
                ログイン済み
              </div>
              <button className="btn btn-ghost" onClick={callApi}>
                更新
              </button>
            </>
          ) : (
            <button className="btn btn-primary" onClick={login}>
              Microsoft でログイン
            </button>
          )}
        </div>
      </header>

      <main className="app-main">
        {apiResult ? (
          <>
            <div className="result-header">
              <span className="result-count">{apiResult.length} 件のメール</span>
            </div>
            <ArrayContainer items={apiResult} />
          </>
        ) : (
          <div className="empty-state">
            <div className="empty-icon">✉</div>
            <div className="empty-title">メールの要約はありません</div>
            <div className="empty-body">
              ログインして「更新」を押すと、受信メールの要約が表示されます。
            </div>
            {!isLoggedIn && (
              <button className="btn btn-primary" style={{ marginTop: 8 }} onClick={login}>
                Microsoft でログイン
              </button>
            )}
          </div>
        )}
      </main>
    </>
  );
}
