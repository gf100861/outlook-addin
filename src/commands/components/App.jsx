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

  // ✅ Corrected version
  // ✅ New function specifically for calling AbstractAPI
  const validateWithAbstractAPI = async (email) => {
    // AbstractAPI URL and your Key
    const API_URL_BASE = "https://emailvalidation.abstractapi.com/v1/";
    const API_KEY = "1b52d865f108441fac2f528e8b925218"; // Your provided sample key

    // 1. Concatenate email and api_key into the full request URL
    const fullUrl = `${API_URL_BASE}?api_key=${API_KEY}&email=${encodeURIComponent(email)}`;

    console.log(`Sending request to AbstractAPI for ${email}: ${fullUrl}`); // Log the request URL

    try {
      // 2. Send GET request, no headers or body needed
      const response = await fetch(fullUrl, {
        method: 'GET', // The method is GET
      });

      const result = await response.json();
      
      console.log("Full response from AbstractAPI for:", email, result); // Log the full API response

      if (!response.ok) {
        console.error("AbstractAPI request failed, status code:", response.status, "Error message:", result);
        // API request failed, e.g., quota exceeded or incorrect key
        return { valid: false, reason: result.error?.message || "API Error" };
      }
      
      // 3. Determine the result based on AbstractAPI's response
      // DELIVERABLE is the most reliable "valid" state
      const isValid = result.deliverability === "DELIVERABLE";

      // AbstractAPI can also provide spelling suggestions!
      const suggestion = result.autocorrect || "";
      
      const finalResult = {
        valid: isValid,
        reason: result.deliverability, // Use the actual deliverability status as the reason
        suggestion: suggestion // Return the spelling suggestion
      };
      
      console.log(`Validation result for ${email}:`, finalResult); // Log the processed result
      
      return finalResult;

    } catch (error) {
      console.error(`Network or other error for ${email}:`, error); // Log network or other errors
      return { valid: false, reason: "Network Error", suggestion: "" };
    }
  };


  const validateEmails = async () => {
    setChecking(true);
    setHasChecked(false);
    setInvalidEmails([]);
    setSuggestedCorrections([]); // Clear old suggestions

    // ... code to get to, cc, bcc lists remains unchanged ...
    const item = Office.context.mailbox.item;
    const to = await fetchEmails(item.to);
    const cc = await fetchEmails(item.cc);
    const bcc = await fetchEmails(item.bcc);

    const allEmails = [...new Set([...to, ...cc, ...bcc])]; // Merge and deduplicate

    // ... preview mode code remains unchanged...
    if (previewOnly) {
      setPreviewData({ to, cc, bcc });
      setChecking(false);
      return;
    }

    const invalid = [];
    const corrections = [];

    // Call the new AbstractAPI validation function
    const validationPromises = allEmails.map(email => validateWithAbstractAPI(email));
    const results = await Promise.all(validationPromises);

    results.forEach((result, index) => {
      if (!result.valid) {
        invalid.push(allEmails[index]);
      }
      // If there is a spelling suggestion, collect it
      if (result.suggestion) {
        corrections.push({
          original: allEmails[index],
          suggested: result.suggestion
        });
      }
    });

    setInvalidEmails(invalid);
    setSuggestedCorrections(corrections); // ✅ Now you can update spelling suggestions!
    setChecking(false);
    setHasChecked(true);
  };
  return (
    <div className={styles.root}>
      <Header logo="assets/Brandlogo.png" title={title} message="Welcome" />
      <HeroList message="Discover what this add-in can do for you today!" items={listItems} />
      {/* Button area (horizontally aligned) */}
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
          {checking ? "Processing..." : previewOnly ? "Preview Emails" : "Validate Emails"}
        </Button>

        <Button appearance="secondary" onClick={() => setPreviewOnly(!previewOnly)}>
          {previewOnly ? "Go to Validation" : "Go to Preview"}
        </Button>
      </div>

      {/* Display area (below buttons) */}
      <div style={{ width: "100%" }}>
        {/* Preview mode display */}
        {previewOnly && (previewData.to.length > 0 || previewData.cc.length > 0 || previewData.bcc.length > 0) && (
          <div
            style={{
              marginTop: "16px",
              padding: "12px",
              border: "1px solid #ccc",
              borderRadius: "6px",
              backgroundColor: "#fafafa",
            }}
          >
            <h4>📧 Recipient Email Preview:</h4>
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

        {/* Suggested email corrections display */}
        {hasChecked && suggestedCorrections.length > 0 && (
          <div style={{ marginTop: "16px", color: "#555" }}>
            <h4>📬 Suggested Email Corrections:</h4>
            <ul>
              {suggestedCorrections.map((item, index) => (
                <li key={index}>
                  Suggest changing <strong>{item.original}</strong> to{" "}
                  <strong style={{ color: "#0066cc" }}>{item.suggested}</strong>
                </li>
              ))}
            </ul>
          </div>
        )}

        {/* Invalid emails display */}
        {hasChecked && invalidEmails.length > 0 && (
          <div style={{ marginTop: "12px" }}>
            <h4 style={{ color: "red" }}>⚠️ The following emails may be invalid:</h4>
            <ul>
              {invalidEmails.map((email, i) => (
                <li key={i}>{email}</li>
              ))}
            </ul>
          </div>
        )}

        {/* All emails passed message */}
        {hasChecked && invalidEmails.length === 0 && !checking && (
          <p style={{ color: "green", marginTop: "12px" }}>✅ All emails passed validation!</p>
        )}
      </div>

    </div>
  );
};

App.propTypes = {
  title: PropTypes.string,
};

export default App;