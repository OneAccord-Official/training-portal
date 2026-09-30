/**
 * POST /.netlify/functions/record-completion
 * body: { trainingId, lessonId, personEmail, score? }
 *
 * Writes one row to the SharePoint "Training Completions" list via Graph.
 * The list's "Person" column is a true Person/Group field, filled in later
 * by a Power Automate flow (trigger: item created here) that resolves
 * PersonEmail to a SharePoint identity — this function only ever writes the
 * plain-text PersonEmail column, which keeps the Graph app registration to
 * plain Sites permissions (no extra SharePoint REST/user-resolution scope).
 */

const { getAccessToken, getSiteId, graphFetch, requireEnv } = require("./_graph");

exports.handler = async (event) => {
  if (event.httpMethod !== "POST") {
    return { statusCode: 405, body: "Method Not Allowed" };
  }

  let payload;
  try {
    payload = JSON.parse(event.body || "{}");
  } catch {
    return { statusCode: 400, body: JSON.stringify({ error: "Invalid JSON body" }) };
  }

  const { trainingId, lessonId, personEmail, score } = payload;
  if (!trainingId || !lessonId || !personEmail) {
    return {
      statusCode: 400,
      body: JSON.stringify({ error: "trainingId, lessonId and personEmail are required" }),
    };
  }

  try {
    const listId = requireEnv("SP_COMPLETIONS_LIST_ID");
    const token = await getAccessToken();
    const siteId = await getSiteId(token);

    const fields = {
      Title: `${trainingId}/${lessonId}`,
      TrainingId: trainingId,
      LessonId: lessonId,
      PersonEmail: personEmail,
      CompletedOn: new Date().toISOString(),
    };
    if (typeof score === "number") fields.Score = score;

    const created = await graphFetch(token, `/sites/${siteId}/lists/${listId}/items`, {
      method: "POST",
      body: JSON.stringify({ fields }),
    });

    return {
      statusCode: 200,
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ ok: true, id: created.id }),
    };
  } catch (err) {
    console.error(err);
    return {
      statusCode: 500,
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ error: err.message }),
    };
  }
};
