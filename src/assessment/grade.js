const REVERSED = new Set(['insufficient_contributions', 'original_unsuitable_contributions']);

// Equal weights; reverse negative criteria. Map the complete 1–5 scale to 1–6.
function suggestedGrade(criteria, answers) {
  if (!criteria?.length) return null;
  const values = criteria.map(c => Number(answers?.[c.id]));
  if (values.some(v => !Number.isInteger(v) || v < 1 || v > 5)) return null;
  const mean = values.reduce((sum, v, i) => sum + (REVERSED.has(criteria[i].id) ? 6 - v : v), 0) / values.length;
  return Math.round((1 + (mean - 1) * 5 / 4) * 10) / 10;
}
function validGrade(value) { return typeof value === 'number' && Number.isFinite(value) && value >= 1 && value <= 6; }
module.exports = { suggestedGrade, validGrade };
