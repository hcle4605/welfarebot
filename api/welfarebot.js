import crypto from "node:crypto";

export const config = { api: { bodyParser: false } }; // HMAC 위해 raw body 필요

export default async function handler(req, res) {
  if (req.method !== "POST") {
    return res.status(405).json({ error: "Method not allowed" });
  }

  // 1) Raw body 읽기
  const rawBody = await readRaw(req);

  // 2) HMAC 검증 (Teams Outgoing Webhook 보안 토큰 기반)
  const auth = req.headers["authorization"] || "";
  const provided = auth.replace(/^HMAC\s*/i, "").trim();
  const signingKeyB64 = process.env.TEAMS_WEBHOOK_SECRET;
  if (!signingKeyB64) return res.status(500).json({ error: "Missing TEAMS_WEBHOOK_SECRET" });
  const calculated = hmacSha256Base64(signingKeyB64, rawBody);
  if (provided !== calculated) {
    return res.status(401).json({ error: "Invalid HMAC signature" });
  }

  // 3) 메시지 파싱
  let activity;
  try { activity = JSON.parse(rawBody); } catch (e) { return res.status(400).json({ error: "Bad JSON" }); }
  const userText = cleanupText(activity?.text || "");
  if (!userText) return res.json({ type: "message", text: "검색할 키워드를 입력해 주세요." });

  // 4) Microsoft Graph 토큰 발급 (Client Credentials)
  const accessToken = await getGraphToken();
  if (!accessToken) return res.status(500).json({ type: "message", text: "Graph 인증 실패" });

  // 5) Search API로 OneDrive/SharePoint 파일 검색
  const hits = await searchDriveItems(accessToken, userText);
  if (!hits.length) {
    return res.json({ type: "message", text: "죄송하지만 관련된 정보를 찾을 수 없습니다. 자세한 사항은 EX팀으로 문의해 주세요." });
  }

  // 6) 요약 (FAST 모드: Graph 스니펫 사용 / FULL 모드: Hugging Face)
  const fast = (process.env.FAST_MODE || "true").toLowerCase() === "true";
  const results = [];
  for (const h of hits.slice(0, 3)) {
    const name = h?.resource?.name || "(제목 없음)";
    const url = h?.resource?.webUrl;
    let snippet = stripHtml(h?.summary || "");

    if (!fast && process.env.HF_API_TOKEN && snippet) {
      try {
        snippet = await summarize(snippet);
      } catch (e) {
        // 실패 시 원본 스니펫 사용
      }
    }
    results.push({ name, url, snippet });
  }

  // 7) Adaptive Card 응답 (5초 제한 내 반환)
  const card = buildAdaptiveCard(userText, results);
  return res.json({ type: "message", attachments: [{ contentType: "application/vnd.microsoft.card.adaptive", content: card }] });
}

// ===================== 유틸 함수들 ===================== //
function readRaw(req) {
  return new Promise((resolve, reject) => {
    let data = "";
    req.setEncoding("utf8");
    req.on("data", chunk => (data += chunk));
    req.on("end", () => resolve(data));
    req.on("error", reject);
  });
}

function hmacSha256Base64(signingKeyB64, bodyText) {
  const key = Buffer.from(signingKeyB64, "base64");
  const h = crypto.createHmac("sha256", key);
  h.update(bodyText, "utf8");
  return h.digest("base64");
}

function cleanupText(t) {
  return (t || "")
    .replace(/^@[^\s]+/g, "") // @복리봇 제거
    .replace(/[\r\n]+/g, " ")
    .trim()
    .slice(0, 200);
}

async function getGraphToken() {
  const tenant = process.env.TENANT_ID;
  const clientId = process.env.CLIENT_ID;
  const clientSecret = process.env.CLIENT_SECRET;
  const body = new URLSearchParams({
    client_id: clientId,
    client_secret: clientSecret,
    grant_type: "client_credentials",
    scope: "https://graph.microsoft.com/.default"
  });
  const resp = await fetch(`https://login.microsoftonline.com/${tenant}/oauth2/v2.0/token`, {
    method: "POST",
    headers: { "Content-Type": "application/x-www-form-urlencoded" },
    body
  });
  if (!resp.ok) return null;
  const json = await resp.json();
  return json.access_token;
}

async function searchDriveItems(token, query) {
  const payload = {
    requests: [
      {
        entityTypes: ["driveItem"],
        region: "APAC",
        query: { queryString: query },
        from: 0,
        size: 5,
        fields: ["name", "webUrl", "fileExtension", "lastModifiedDateTime", "size"]
      }
    ]
  };

  const r = await fetch("https://graph.microsoft.com/v1.0/search/query", {
    method: "POST",
    headers: { Authorization: `Bearer ${token}`, "Content-Type": "application/json" },
    body: JSON.stringify(payload)
  });
  if (!r.ok) return [];
  const data = await r.json();
  return data?.value?.[0]?.hitsContainers?.[0]?.hits || [];
}

function stripHtml(html) {
  return (html || "")
    .replace(/<[^>]+>/g, " ")
    .replace(/&[^;]+;/g, " ")
    .replace(/[\s\u00A0]+/g, " ")
    .trim();
}

async function summarize(text) {
  const inp = text.slice(0, 3000);
  const r = await fetch("https://api-inference.huggingface.co/models/facebook/bart-large-cnn", {
    method: "POST",
    headers: {
      Authorization: `Bearer ${process.env.HF_API_TOKEN}`,
      "Content-Type": "application/json"
    },
    body: JSON.stringify({ inputs: inp, parameters: { max_length: 120, min_length: 40 } })
  });
  const out = await r.json();
  const s = Array.isArray(out) ? out[0]?.summary_text : out?.summary_text;
  return (s || inp).trim();
}

function buildAdaptiveCard(query, items) {
  return {
    "type": "AdaptiveCard",
    "$schema": "http://adaptivecards.io/schemas/adaptive-card.json",
    "version": "1.4",
    "body": [
      { "type": "TextBlock", "size": "Medium", "weight": "Bolder", "text": `\uD83D\uDD0D '${query}' 검색 결과` },
      ...items.map(i => ({
        type: "Container",
        items: [
          { type: "TextBlock", wrap: true, weight: "Bolder", text: `[${i.name}](${i.url})` },
          { type: "TextBlock", wrap: true, spacing: "Small", text: i.snippet || "요약 스니펫 없음" }
        ]
      }))
    ]
  };
}