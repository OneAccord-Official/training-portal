/**
 * GET /.netlify/functions/get-completions?personEmail=jane@oneaccord.co
 *
 * Returns the list of { trainingId, lessonId } this person has already
 * completed, so the portal can restore checkmarks in the nav across
 * sessions/devices instead of relying on localStorage.
 */

const { getAccessToken, getSiteId, graphFetch, requireEnv } = require("./_graph");

exports.handler = async (event) => {
  const personEmail = event.queryStringParameters && event.queryStringParameters.personEmail;
  if (!personEmail) {
    return { statusCode: 400, body: JSON.stringify({ error: "personEmail query param is required" }) };
  }

  try {
    const listId = requireEnv("SP_COMPLETIONS_LIST_ID");
    const token = await getAccessToken();
    const siteId = await getSiteId(token);

    // Custom fields aren't indexed by default, so a $filter on them needs
    // this Prefer header; fine at this list's expected scale.
    const safeEmail = personEmail.replace(/'/g, "''");
    const data = await graphFetch(
      token,
      `/sites/${siteId}/lists/${listId}/items?expand=fields(select=TrainingId,LessonId,CompletedOn,Score)&$filter=fields/PersonEmail eq '${safeEmail}'&$top=500`,
      { headers: { Prefer: "HonorNonIndexedQueriesWarningMayFailRandomly" } }
    );

    const completions = (data.value || [])
      .map((item) => item.fields)
      .filter((f) => f && f.TrainingId && f.LessonId)
      .map((f) => ({
        trainingId: f.TrainingId,
        lessonId: f.LessonId,
        completedOn: f.CompletedOn,
        score: f.Score ?? null,
      }));

    return {
      statusCode: 200,
      headers: { "Content-Type": "application/json", "Cache-Control": "no-store" },
      body: JSON.stringify({ completions }),
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
