import * as React from "react";
import PropTypes from "prop-types";
import { Image, tokens, makeStyles } from "@fluentui/react-components";

const useStyles = makeStyles({
  header: {
    display: "flex",
    flexDirection: "column",
    alignItems: "center",
    justifyContent: "center",
    paddingTop: "48px",
    paddingBottom: "32px",
    paddingLeft: "24px",
    paddingRight: "24px",
    backgroundColor: tokens.colorNeutralBackground1,
    borderBottom: `1px solid ${tokens.colorNeutralStroke2}`,
    boxShadow: tokens.shadow16,
    borderRadius: "0 0 16px 16px",
  },
  logo: {
    width: "72px",
    height: "72px",
    marginBottom: "16px",
    borderRadius: "12px",
    boxShadow: tokens.shadow8,
  },
  title: {
    fontSize: tokens.fontSizeHero800,
    fontWeight: tokens.fontWeightSemibold,
    color: tokens.colorBrandForeground1,
    margin: 0,
    textAlign: "center",
    lineHeight: "1.2",
  },
  subtitle: {
    fontSize: tokens.fontSizeBase500,
    color: tokens.colorNeutralForeground3,
    marginTop: "8px",
    textAlign: "center",
    maxWidth: "280px",
    lineHeight: "1.4",
  },
});

const Header = ({ title, logo, message }) => {
  const styles = useStyles();

  return (
    <section className={styles.header}>
      <Image src={logo} alt={title} className={styles.logo} />
      <h1 className={styles.title}>{message}</h1>
      {title && <p className={styles.subtitle}>{title}</p>}
    </section>
  );
};

Header.propTypes = {
  title: PropTypes.string, // 副标题
  logo: PropTypes.string,
  message: PropTypes.string, // 主标题
};

export default Header;
