// Formula order and rounding follow Ireland-Import-Cost-Calculator.xlsx, F6:G15.
export const IRELAND_DEFAULTS = Object.freeze({exchange: 1.163467, duty: 2.7, vat: 23, processing: 2.5, minimum: 14.5, lower: 5, upper: 20, smallFee: 5});
const cents = value => Math.round((value + Number.EPSILON * value) * 100) / 100;
export function calculateIrelandDelivery(input) {
  const fields = ['goods', 'delivery', 'insurance', ...Object.keys(IRELAND_DEFAULTS)];
  const values = Object.fromEntries(fields.map(key => [key, input[key] === '' || input[key] == null ? NaN : Number(input[key])]));
  if (fields.some(key => !Number.isFinite(values[key]) || values[key] < 0 || values[key] > 1e12)) return {error: 'Enter a valid amount in each field. Use 0 if there is no cost.'};
  if (values.exchange <= 0) return {error: 'The exchange rate must be greater than 0.'};
  if (values.upper < values.lower) return {error: 'The upper fee threshold must be at least the lower threshold.'};
  const {goods, delivery, insurance, exchange, duty, vat, processing, minimum, lower, upper, smallFee} = values;
  const customs = cents((goods + delivery + insurance) * exchange);
  const dutyCost = cents(customs * duty / 100);
  const vatCost = cents((customs + dutyCost) * vat / 100);
  const processingCost = cents(customs < lower ? 0 : customs <= upper ? smallFee : Math.max(minimum, customs * processing / 100));
  const additional = cents(dutyCost + vatCost + processingCost);
  const deliveryTotal = cents((delivery + insurance) * exchange + additional);
  const landed = cents(customs + additional);
  const eur = {customs, duty: dutyCost, vat: vatCost, processing: processingCost, additional, deliveryTotal, landed};
  return {eur, gbp: Object.fromEntries(Object.entries(eur).map(([key, amount]) => [key, cents(amount / exchange)]))};
}
