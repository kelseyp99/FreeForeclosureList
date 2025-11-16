import React, { useState } from "react";
import { signInWithGoogle, signOutGoogle, auth } from "./firebase";
import { onAuthStateChanged } from "firebase/auth";

export default function GoogleAuthButton() {
  const [user, setUser] = useState(null);

  React.useEffect(() => {
    const unsubscribe = onAuthStateChanged(auth, (u) => setUser(u));
    return () => unsubscribe();
  }, []);

  if (user) {
    return (
      <div style={{ margin: '16px 0' }}>
        <span style={{ marginRight: 12 }}>Signed in as {user.displayName}</span>
        <button onClick={signOutGoogle}>Sign Out</button>
      </div>
    );
  }
  return (
    <button onClick={signInWithGoogle} style={{ margin: '16px 0' }}>
      Sign in with Google
    </button>
  );
}
