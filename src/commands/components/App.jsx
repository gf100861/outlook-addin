import * as React from "react";
import PropTypes from "prop-types";
import Header from "./Header";
import HeroList from "./HeroList";
import { makeStyles, Button } from "@fluentui/react-components";
import {
  Mail24Regular,
  ShieldCheckmark24Regular,
  Lightbulb24Regular,
} from "@fluentui/react-icons";

const useStyles = makeStyles({
  root: {
    minHeight: "100vh",
    padding: "16px",
  },
});

const App = (props) => {
  const { title } = props;
  const styles = useStyles();

  const [invalidEmails, setInvalidEmails] = React.useState([]);
  const [checking, setChecking] = React.useState(false);
  const [hasChecked, setHasChecked] = React.useState(false);
  const [suggestedCorrections, setSuggestedCorrections] = React.useState([]);
  const [previewOnly, setPreviewOnly] = React.useState(true);
  const [previewData, setPreviewData] = React.useState({ to: [], cc: [], bcc: [] });

  const listItems = [
    {
      icon: <Mail24Regular />,
      primaryText: "Check email syntax and domain",
    },
    {
      icon: <ShieldCheckmark24Regular />,
      primaryText: "Avoid invalid or risky emails",

    },
    {
      icon: <Lightbulb24Regular />,
      primaryText: "Improve input with live hints",

    },
  ];

  const splitEmails = (raw) =>
    raw
      .split(/[;,]+/)
      .map((e) => e.trim().toLowerCase())
      .filter(Boolean);

  const fetchEmails = (field) =>
    new Promise((resolve) => {
      field.getAsync((res) => {
        if (res.status === Office.AsyncResultStatus.Succeeded) {
          const list = res.value.flatMap((e) => splitEmails(e.emailAddress));
          resolve(list);
        } else {
          resolve([]);
        }
      });
    });

  const delay = (ms) => new Promise((resolve) => setTimeout(resolve, ms));

  // ✅ 修正后的版本
  // ✅ 专门用于调用 AbstractAPI 的新函数
  const validateWithAbstractAPI = async (email) => {
    // AbstractAPI 的 URL 和你的 Key
    const API_URL_BASE = "https://emailvalidation.abstractapi.com/v1/";
    const API_KEY = "1b52d865f108441fac2f528e8b925218"; // 这是你提供的示例 Key

    // 1. 将 email 和 api_key 拼接成完整的请求 URL
    const fullUrl = `${API_URL_BASE}?api_key=${API_KEY}&email=${encodeURIComponent(email)}`;

    try {
      // 2. 发送 GET 请求，不需要 headers 和 body
      const response = await fetch(fullUrl, {
        method: 'GET', // 方法是 GET
      });

      const result = await response.json();

      if (!response.ok) {
        console.error("AbstractAPI 请求失败，状态码:", response.status, "错误信息:", result);
        // API 请求失败，例如超出配额或 key 错误
        return { valid: false, reason: result.error?.message || "API Error" };
      }

      console.log("AbstractAPI 验证结果:", email, result);

      // 3. 根据 AbstractAPI 的响应来判断结果
      // DELIVERABLE 是最可靠的“有效”状态
      const isValid = result.deliverability === "DELIVERABLE";

      // AbstractAPI 还能提供拼写建议！
      const suggestion = result.autocorrect || "";

      return {
        valid: isValid,
        reason: result.deliverability, // 将真实的 deliverability 状态作为原因
        suggestion: suggestion // 返回拼写建议
      };

    } catch (error) {
      console.error("网络或其他错误:", email, error);
      return { valid: false, reason: "Network Error", suggestion: "" };
    }
  };


  const validateEmails = async () => {
    setChecking(true);
    setHasChecked(false);
    setInvalidEmails([]);
    setSuggestedCorrections([]); // 清空旧的建议

    // ... 获取 to, cc, bcc 列表的代码保持不变 ...
    const item = Office.context.mailbox.item;
    const to = await fetchEmails(item.to);
    const cc = await fetchEmails(item.cc);
    const bcc = await fetchEmails(item.bcc);

    const allEmails = [...new Set([...to, ...cc, ...bcc])]; // 合并并去重

    // ...预览模式的代码不变...
    if (previewOnly) {
      // ...
      return;
    }

    const invalid = [];
    const corrections = [];

    // 调用新的 AbstractAPI 验证函数
    const validationPromises = allEmails.map(email => validateWithAbstractAPI(email));
    const results = await Promise.all(validationPromises);

    results.forEach((result, index) => {
      if (!result.valid) {
        invalid.push(allEmails[index]);
      }
      // 如果有拼写建议，就收集起来
      if (result.suggestion) {
        corrections.push({
          original: allEmails[index],
          suggested: result.suggestion
        });
      }
    });

    setInvalidEmails(invalid);
    setSuggestedCorrections(corrections); // ✅ 现在可以更新拼写建议了！
    setChecking(false);
    setHasChecked(true);
  };
  return (
    <div className={styles.root}>
      <Header logo="assets/Brandlogo.png" title={title} message="Welcome" />
      <HeroList message="Discover what this add-in can do for you today!" items={listItems} />
      {/* 按钮区（横向对齐） */}
      <div
        style={{
          display: "flex",
          justifyContent: "space-between",
          alignItems: "center",
          width: "100%",
          marginTop: "16px",
        }}
      >
        <Button appearance="primary" onClick={validateEmails} disabled={checking}>
          {checking ? "正在处理..." : previewOnly ? "预览邮箱" : "验证邮箱"}
        </Button>

        <Button appearance="secondary" onClick={() => setPreviewOnly(!previewOnly)}>
          {previewOnly ? "进入验证" : "进入预览"}
        </Button>
      </div>

      {/* 展示区（放按钮下方） */}
      <div style={{ width: "100%" }}>
        {/* 预览模式展示 */}
        {previewOnly && (previewData.to.length || previewData.cc.length || previewData.bcc.length) > 0 && (
          <div
            style={{
              marginTop: "16px",
              padding: "12px",
              border: "1px solid #ccc",
              borderRadius: "6px",
              backgroundColor: "#fafafa",
            }}
          >
            <h4>📧 收件人邮箱预览：</h4>
            {previewData.to.length > 0 && (
              <>
                <strong>To:</strong>
                <ul>{previewData.to.map((email, i) => <li key={`to-${i}`}>{email}</li>)}</ul>
              </>
            )}
            {previewData.cc.length > 0 && (
              <>
                <strong>Cc:</strong>
                <ul>{previewData.cc.map((email, i) => <li key={`cc-${i}`}>{email}</li>)}</ul>
              </>
            )}
            {previewData.bcc.length > 0 && (
              <>
                <strong>Bcc:</strong>
                <ul>{previewData.bcc.map((email, i) => <li key={`bcc-${i}`}>{email}</li>)}</ul>
              </>
            )}
          </div>
        )}

        {/* 建议邮箱修正展示 */}
        {hasChecked && suggestedCorrections.length > 0 && (
          <div style={{ marginTop: "16px", color: "#555" }}>
            <h4>📬 建议邮箱修正：</h4>
            <ul>
              {suggestedCorrections.map((item, index) => (
                <li key={index}>
                  建议将 <strong>{item.original}</strong> 修改为{" "}
                  <strong style={{ color: "#0066cc" }}>{item.suggested}</strong>
                </li>
              ))}
            </ul>
          </div>
        )}

        {/* 无效邮箱展示 */}
        {hasChecked && invalidEmails.length > 0 && (
          <div style={{ marginTop: "12px" }}>
            <h4 style={{ color: "red" }}>⚠️ 以下邮箱可能无效：</h4>
            <ul>
              {invalidEmails.map((email, i) => (
                <li key={i}>{email}</li>
              ))}
            </ul>
          </div>
        )}

        {/* 所有邮箱通过提示 */}
        {hasChecked && invalidEmails.length === 0 && !checking && (
          <p style={{ color: "green", marginTop: "12px" }}>✅ 所有邮箱验证通过！</p>
        )}
      </div>

    </div>
  );
};

App.propTypes = {
  title: PropTypes.string,
};

export default App;
