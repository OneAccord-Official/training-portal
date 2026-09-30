/**
 * GET /.netlify/functions/get-trainings
 *
 * Reads the SharePoint "Trainings" list and reshapes it into the same
 * { trainings: [ { id, title, lessons: [...] } ] } structure the portal's
 * nav rendering already expects (previously hand-written in
 * data/trainings-index.json). Grouping key is TrainingId; each row becomes
 * one lesson entry, sorted by Order.
 */

const { getAccessToken, getSiteId, graphFetch, requireEnv } = require("./_graph");

exports.handler = async () => {
  try {
    const listId = requireEnv("SP_TRAININGS_LIST_ID");
    const token = await getAccessToken();
    const siteId = await getSiteId(token);

    const data = await graphFetch(
      token,
      `/sites/${siteId}/lists/${listId}/items?expand=fields(select=Title,TrainingId,LessonId,Category,Order,Status,ContentUrl)&$top=500`
    );

    const rows = (data.value || [])
      .map((item) => item.fields)
      .filter((f) => f && f.TrainingId && f.LessonId && f.Status !== "Retired");

    const byTraining = new Map();
    for (const row of rows) {
      if (!byTraining.has(row.TrainingId)) {
        byTraining.set(row.TrainingId, {
          id: row.TrainingId,
          title: row.Category || row.TrainingId,
          lessons: [],
        });
      }
      byTraining.get(row.TrainingId).lessons.push({
        id: row.LessonId,
        title: row.Title || row.LessonId,
        status: (row.Status || "Pending").toLowerCase(),
        order: typeof row.Order === "number" ? row.Order : 999,
        contentUrl: row.ContentUrl || `data/trainings/${row.TrainingId}/${row.LessonId}.json`,
      });
    }

    const trainings = Array.from(byTraining.values()).map((t) => ({
      ...t,
      lessons: t.lessons.sort((a, b) => a.order - b.order),
    }));

    return {
      statusCode: 200,
      headers: { "Content-Type": "application/json", "Cache-Control": "no-store" },
      body: JSON.stringify({ trainings }),
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
