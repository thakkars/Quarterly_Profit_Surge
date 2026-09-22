export async function onRequest(context) {
  // 🛑 CHANGE YOUR USERNAME AND PASSWORD HERE 🛑
  const expectedUsername = "admin"; 
  const expectedPassword = "Pass@123"; 

  const authHeader = context.request.headers.get("Authorization");

  if (!authHeader) {
    return new Response("Unauthorized", {
      status: 401,
      headers: { "WWW-Authenticate": 'Basic realm="Secure Screener Area"' },
    });
  }

  const [scheme, encoded] = authHeader.split(" ");
  if (scheme !== "Basic" || !encoded) {
    return new Response("Malformed authorization header", { status: 400 });
  }

  const decoded = atob(encoded);
  const [username, password] = decoded.split(":");

  if (username === expectedUsername && password === expectedPassword) {
    return context.next();
  }

  return new Response("Unauthorized. Wrong username or password.", {
    status: 401,
    headers: { "WWW-Authenticate": 'Basic realm="Secure Screener Area"' },
  });
}
