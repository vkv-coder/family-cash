// Relays a Telegram message using the bot token stored as a Supabase secret,
// so the token never ships to the browser. Deploy with:
//   supabase functions deploy send-telegram
// then set the secret once with:
//   supabase secrets set TG_BOT_TOKEN=<new-token>

const corsHeaders = {
  "Access-Control-Allow-Origin": "*",
  "Access-Control-Allow-Headers": "authorization, x-client-info, apikey, content-type",
  "Access-Control-Allow-Methods": "POST, OPTIONS",
};

Deno.serve(async (req: Request) => {
  if (req.method === "OPTIONS") {
    return new Response("ok", { headers: corsHeaders });
  }

  if (req.method !== "POST") {
    return new Response("Method not allowed", { status: 405, headers: corsHeaders });
  }

  const botToken = Deno.env.get("TG_BOT_TOKEN");
  if (!botToken) {
    return new Response(JSON.stringify({ error: "TG_BOT_TOKEN not configured" }), {
      status: 500,
      headers: { ...corsHeaders, "Content-Type": "application/json" },
    });
  }

  let chat_id: string | undefined;
  let text: string | undefined;
  try {
    ({ chat_id, text } = await req.json());
  } catch {
    return new Response(JSON.stringify({ error: "Invalid JSON body" }), {
      status: 400,
      headers: { ...corsHeaders, "Content-Type": "application/json" },
    });
  }

  if (!chat_id || !text) {
    return new Response(JSON.stringify({ error: "chat_id and text are required" }), {
      status: 400,
      headers: { ...corsHeaders, "Content-Type": "application/json" },
    });
  }

  const tgRes = await fetch(`https://api.telegram.org/bot${botToken}/sendMessage`, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ chat_id, text }),
  });

  const tgData = await tgRes.json();
  return new Response(JSON.stringify(tgData), {
    status: tgRes.ok ? 200 : 502,
    headers: { ...corsHeaders, "Content-Type": "application/json" },
  });
});
