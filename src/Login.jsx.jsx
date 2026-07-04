import { useState } from "react";

export default function Login() {
  const [email, setEmail] = useState("");
  const [code, setCode] = useState("");
  const [step, setStep] = useState("email");

  const requestCode = async () => {
    await fetch("https://auth.abapps1.co.uk/auth/request-code", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ email })
    });

    setStep("code");
  };

  const verifyCode = async () => {
    const res = await fetch("https://auth.abapps1.co.uk/auth/verify-code", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ email, code })
    });

    const data = await res.json();

    if (data.token) {
      localStorage.setItem("token", data.token);
      alert("Logged in!");
    } else {
      alert("Invalid code");
    }
  };

  return (
    <div style={{ padding: 20 }}>
      {step === "email" && (
        <>
          <h2>Login</h2>
          <input
            type="email"
            placeholder="Enter your email"
            value={email}
            onChange={(e) => setEmail(e.target.value)}
          />
          <button onClick={requestCode}>Send Code</button>
        </>
      )}

      {step === "code" && (
        <>
          <h2>Enter Code</h2>
          <input
            type="text"
            placeholder="6-digit code"
            value={code}
            onChange={(e) => setCode(e.target.value)}
          />
          <button onClick={verifyCode}>Verify</button>
        </>
      )}
    </div>
  );
}
