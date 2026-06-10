import React, { useEffect, useState } from "react";
import { PublicClientApplication } from "@azure/msal-browser";

const msalInstance = new PublicClientApplication({
  auth: {
    clientId: "d843a25b-962d-4912-9740-5d7f83ede221",
    authority: "https://login.microsoftonline.com/a7929d4b-ae1e-4e78-8ff8-3293161eac3a",
    redirectUri: "http://localhost:3000"
  }
});

export default function App() {

  const [isLoggedIn, setIsLoggedIn] = useState(false);
  const [apiResult, setApiResult] = useState(null);

  // ✅ ログイン後の処理
  useEffect(() => {
    const init = async () => {
      await msalInstance.initialize();

      const res = await msalInstance.handleRedirectPromise();

      if (res) {
        console.log("ログイン成功", res.idToken);
        setIsLoggedIn(true); // ✅ 表示用
      }

      // 既にログイン済みの場合も対応
      const accounts = msalInstance.getAllAccounts();
      if (accounts.length > 0) {
        setIsLoggedIn(true);
      }
    };

    init();
  }, []);

  // ✅ ログイン
  const login = async () => {
    await msalInstance.loginRedirect({
      scopes: ["openid", "profile", "email"]
    });
  };

  // ✅ API呼び出し
  const callApi = async () => {
    const accounts = msalInstance.getAllAccounts();

    if (accounts.length === 0) {
      alert("ログインしていません");
      return;
    }

    const tokenRes = await msalInstance.acquireTokenSilent({
      account: accounts[0],
      scopes: ["openid", "profile", "email"]
    });

    const token = tokenRes.idToken;

    const res = await fetch(
      "https://pd3bdzc1wg.execute-api.ap-northeast-1.amazonaws.com/summaries",
      {
        headers: {
          Authorization: `Bearer ${token}`
        }
      }
    );

    const data = await res.json();

    console.log("API:", data);

    setApiResult(data); // ✅ 表示用
  };

  return (
    <div style={{ padding: 40 }}>
      <h1>JWT テスト</h1>

      {/* ✅ ログインボタン */}
      {!isLoggedIn && (
        <button onClick={login}>ログイン</button>
      )}

      {/* ✅ ログイン成功表示 */}
      {isLoggedIn && (
        <p style={{ color: "green" }}>
          ✅ ログイン成功
        </p>
      )}

      {/* ✅ APIボタン */}
      {isLoggedIn && (
        <button onClick={callApi}>
          API呼び出し
        </button>
      )}

      {/* ✅ API結果表示 */}
      {apiResult && (
        <div style={{ marginTop: 20 }}>
          <h3>APIレスポンス</h3>
          <pre>{JSON.stringify(apiResult, null, 2)}</pre>
        </div>
      )}

    </div>
  );
}