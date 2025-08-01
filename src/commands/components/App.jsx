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

    const validateEmailWithOwnBackend = async (email) => {
    // 您的后端服务器地址和API密钥
    const API_URL = "https://email-validator-backend.vercel.app/api/validate";
    const API_KEY = "hj122400";

    try {
      const response = await fetch(API_URL, {
        method: 'POST',
        headers: {
          'Content-Type': 'application/json',
          'X-API-KEY': API_KEY
        },
        body: JSON.stringify({ email: email })
      });

      if (!response.ok) {
        console.error("API request failed with status:", response.status);
        // 如果API请求失败 (例如密钥错误), 我们默认它无效
        return { valid: false };
      }

      const result = await response.json();
      console.log("后端验证结果:", email, result);
      
      // 根据您的API响应来返回结果
      return { valid: result.is_valid };

    } catch (error) {
      console.error("验证失败:", email, error);
      // 网络错误等情况，我们也默认它无效
      return { valid: false };
    }
  };


  const validateEmails = async () => {
    setChecking(true);
    setHasChecked(false);
    setInvalidEmails([]); // Clear previous results
    setSuggestedCorrections([]); // Clear previous results

    // ... 获取 to, cc, bcc 列表的代码保持不变 ...
    const item = Office.context.mailbox.item;
    const to = await fetchEmails(item.to);
    const cc = await fetchEmails(item.cc);
    const bcc = await fetchEmails(item.bcc);

    const toList = [...new Set(to)];
    const ccList = [...new Set(cc)];
    const bccList = [...new Set(bcc)];
    // ...

    if (previewOnly) {
      setPreviewData({ to: toList, cc: ccList, bcc: bccList });
      setChecking(false);
      return;
    }

    const allEmails = [...new Set([...toList, ...ccList, ...bccList])];
    const invalid = [];

    // 使用 Promise.all 来并行处理所有API请求，以提高速度
    const validationPromises = allEmails.map(email => validateEmailWithOwnBackend(email));
    const results = await Promise.all(validationPromises);

    results.forEach((result, index) => {
      if (!result.valid) {
        invalid.push(allEmails[index]);
      }
    });

    setInvalidEmails(invalid);
    // 注意：这个后端版本没有拼写建议功能，您可以选择在前端保留或移除
    // setSuggestedCorrections(corrections); 
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
