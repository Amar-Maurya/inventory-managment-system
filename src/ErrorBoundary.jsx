import React from "react";

// Catches any error thrown during render anywhere below it in the tree.
// Without this, a single component crash unmounts the entire app and
// leaves a blank white page with no way to recover except manually
// navigating away — this gives people a way back in instead.
export default class ErrorBoundary extends React.Component {
  constructor(props) {
    super(props);
    this.state = { hasError: false, error: null };
  }

  static getDerivedStateFromError(error) {
    return { hasError: true, error };
  }

  componentDidCatch(error, info) {
    console.error("GroutLine crashed:", error, info);
  }

  handleReload = () => {
    window.location.reload();
  };

  render() {
    if (this.state.hasError) {
      return (
        <div style={{ fontFamily: "Inter, sans-serif", height: "100vh", display: "flex", alignItems: "center", justifyContent: "center", background: "#FAFAF9", padding: "24px" }}>
          <div style={{ maxWidth: "380px", textAlign: "center" }}>
            <h1 style={{ fontSize: "20px", fontWeight: 600, marginBottom: "8px", color: "#22262A" }}>
              Something went wrong
            </h1>
            <p style={{ fontSize: "14px", color: "#8A9499", marginBottom: "20px", lineHeight: 1.5 }}>
              The app hit an unexpected error. Your data is safe in your Google Sheet — reloading should fix this.
            </p>
            <button
              onClick={this.handleReload}
              style={{ background: "#1C2B33", color: "#fff", border: "none", borderRadius: "8px", padding: "10px 20px", fontSize: "14px", fontWeight: 500, cursor: "pointer" }}
            >
              Reload app
            </button>
          </div>
        </div>
      );
    }
    return this.props.children;
  }
}
