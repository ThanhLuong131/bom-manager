export async function onRequestGet({ request, env }) {
  const url = new URL(request.url);
  const userId = url.searchParams.get("user_id");
  if (!userId) return Response.json({ error: "Missing user_id" }, { status: 400 });
  const { results } = await env.bom_db
    .prepare("SELECT * FROM records WHERE user_id=? ORDER BY created_at DESC")
    .bind(userId)
    .all();
  return Response.json({ records: results });
}

export async function onRequestPost({ request, env }) {
  const body = await request.json();
  const { user_id, part_number, product_name, bom_cost, margin, exchange_rate, qty, date, to_customer, attention, term, tel, remark, notes, file_name } = body;
  const result = await env.bom_db
    .prepare(`INSERT INTO records (user_id,part_number,product_name,bom_cost,margin,exchange_rate,qty,date,to_customer,attention,term,tel,remark,notes,file_name)
              VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)`)
    .bind(user_id, part_number, product_name, bom_cost, margin, exchange_rate, qty||"", date||"", to_customer||"", attention||"", term||"", tel||"", remark||"", notes||"", file_name||"")
    .run();
  return Response.json({ id: result.meta.last_row_id });
}

export async function onRequestPut({ request, env }) {
  const body = await request.json();
  const { id, part_number, product_name, bom_cost, margin, exchange_rate, qty, date, to_customer, attention, term, tel, remark, notes } = body;
  await env.bom_db
    .prepare(`UPDATE records SET part_number=?,product_name=?,bom_cost=?,margin=?,exchange_rate=?,qty=?,date=?,to_customer=?,attention=?,term=?,tel=?,remark=?,notes=? WHERE id=?`)
    .bind(part_number, product_name, bom_cost, margin, exchange_rate, qty||"", date||"", to_customer||"", attention||"", term||"", tel||"", remark||"", notes||"", id)
    .run();
  return Response.json({ ok: true });
}

export async function onRequestDelete({ request, env }) {
  const { id } = await request.json();
  await env.bom_db.prepare("DELETE FROM records WHERE id=?").bind(id).run();
  return Response.json({ ok: true });
}