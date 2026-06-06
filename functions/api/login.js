export async function onRequestPost({ request, env }) {
  const { username, password } = await request.json();
  const user = await env.bom_db
    .prepare("SELECT id, username, name FROM users WHERE username=? AND password=?")
    .bind(username, password)
    .first();
  if (!user) {
    return Response.json({ error: "Sai tài khoản hoặc mật khẩu" }, { status: 401 });
  }
  return Response.json({ user });
}
