import * as React from "react";
import PropTypes from "prop-types";
import Header from "./Header";
import HeroList from "./HeroList";
import { makeStyles, Button } from "@fluentui/react-components";
import {
  Ribbon24Regular,
  LockOpen24Regular,
  DesignIdeas24Regular,
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
      icon: <Ribbon24Regular />,
      primaryText: "Achieve more with Office integration",
    },
    {
      icon: <LockOpen24Regular />,
      primaryText: "Unlock features and functionality",
    },
    {
      icon: <DesignIdeas24Regular />,
      primaryText: "Create and visualize like a pro",
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

  const validateEmailWithResult = async (email) => {
    const API_KEY = "a2efde7285824ec99b3822e4c103bad1";
    const url = `https://emailvalidation.abstractapi.com/v1/?api_key=${API_KEY}&email=${encodeURIComponent(email)}`;

    try {
      const res = await fetch(url);
      const result = await res.json();
      console.log("验证结果:", email, result);

      const {
        is_valid_format,
        is_disposable_email,
        is_mx_found,
        is_smtp_valid,
        deliverability,
        autocorrect,
      } = result;

      const valid =
        is_valid_format?.value === true &&
        is_disposable_email?.value === false &&
        is_mx_found?.value === true &&
        is_smtp_valid?.value === true &&
        deliverability !== "UNKNOWN";

      return { valid, autocorrect };
    } catch (err) {
      console.error("验证失败：", email, err);
      return { valid: false, autocorrect: null };
    }
  };

  const validateEmails = async () => {
    setChecking(true);
    setHasChecked(false);

    const item = Office.context.mailbox.item;
    const to = await fetchEmails(item.to);
    const cc = await fetchEmails(item.cc);
    const bcc = await fetchEmails(item.bcc);

    const toList = [...new Set(to)];
    const ccList = [...new Set(cc)];
    const bccList = [...new Set(bcc)];

    if (previewOnly) {
      setPreviewData({ to: toList, cc: ccList, bcc: bccList });
      setChecking(false);
      return;
    }

    const allEmails = [...new Set([...toList, ...ccList, ...bccList])];
    const invalid = [];
    const corrections = [];

    for (const email of allEmails) {
      const result = await validateEmailWithResult(email);

      if (!result.valid) {
        invalid.push(email);
      }

      if (result.autocorrect && result.autocorrect !== email) {
        corrections.push({ original: email, suggested: result.autocorrect });
      }

      await delay(1000); // 限速：每秒1次，避免 API 限流
    }

    setInvalidEmails(invalid);
    setSuggestedCorrections(corrections);
    setChecking(false);
    setHasChecked(true);
  };

  return (
    <div className={styles.root}>
      <Header logo="assets/logo-filled.png" title={title} message="Welcome" />
      <HeroList message="Discover what this add-in can do for you today!" items={listItems} />

      <div style={{ marginTop: "16px" }}>
        <Button appearance="primary" onClick={validateEmails} disabled={checking}>
          {checking ? "正在处理..." : previewOnly ? "预览收件人邮箱" : "验证收件人邮箱地址"}
        </Button>

        <Button
          appearance="secondary"
          onClick={() => setPreviewOnly(!previewOnly)}
          style={{ marginLeft: "12px" }}
        >
          {previewOnly ? "切换到验证模式" : "切换到预览模式"}
        </Button>

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
            <h4 style={{ color: "red" }}>⚠️ 以下邮箱无效：</h4>
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
