/**
 * =========================================================
 * VAIBHAV's GOOGLE ADS PERFORMANCE EXPORT SCRIPT
 * BY CONVERSION DATE VERSION — 
 * =========================================================
 *
 * EXPORTS:
 * - Campaign Report           (by conversion date)
 * - Ad Group Report           (by conversion date)
 * - Keyword Report
 * - Search Terms              (by click date — API limit)
 * - Product Performance       (by click date — API limit)
 * - Device Report             (by click date — API limit)
 * - Geo Report                (by click date — API limit)
 * - Day of Week Performance   (by click date)
 * - Funnel Analysis           (derived insights)
 *
 * CUSTOM METRICS INCLUDED:
 * - Conv (GA4 Purchase, by conv. time)
 * - Revenue (GA4 Purchase, by conv. time)
 * - ROAS (by conv. time)
 * - CPA (by conv. time)
 * - ATC (by conv. time)
 *
 * IMPORTANT:
 * by-conversion-date metrics only work on:
 *   campaign, ad_group, ad_group_ad
 * They do NOT work on:
 *   search_term_view, shopping_performance_view,
 *   geographic_view, or anything using segments.device
 * Those reports use standard all_conversions /
 * all_conversions_value (click-date attribution).
 *
 * =========================================================
 */

const SPREADSHEET_URL =
  'https://docs.google.com/spreadsheets/d/1E18TdEnAJm0BVrfKGoO7HNo83yjBesdCjQSo1JAeBcA/edit?gid=0#gid=0';

/**
 * DATE RANGE
 *
 * Full March + April 2026 for MoM comparison.
 */
const DATE_FROM = '20260301';
const DATE_TO = '20260430';

/**
 * CONVERSION ACTION NAMES
 *
 * Single GA4 purchase tag for this account.
 * ATC tag is from the Google Shopping App feed.
 */
const CONV = {
  PURCHASE: 'Hunnit (web) purchase',
  ATC: 'Google Shopping App Add To Cart'
};

/**
 * =========================================================
 * MAIN
 * =========================================================
 */

function main() {

  const spreadsheet =
    SpreadsheetApp.openByUrl(
      SPREADSHEET_URL
    );

  exportCampaignReport(spreadsheet);
  exportAdGroupReport(spreadsheet);
  exportKeywordReport(spreadsheet);
  exportSearchTermsReport(spreadsheet);
  exportProductReport(spreadsheet);
  exportDeviceReport(spreadsheet);
  exportGeoReport(spreadsheet);
  exportDayOfWeekReport(spreadsheet);
  exportFunnelAnalysis(spreadsheet);

  Logger.log(
    'DONE - All reports exported successfully.'
  );
}

/**
 * =========================================================
 * HELPERS
 * =========================================================
 */

function getOrCreateSheet(
  spreadsheet,
  sheetName
) {

  let sheet =
    spreadsheet.getSheetByName(sheetName);

  if (!sheet) {

    sheet =
      spreadsheet.insertSheet(sheetName);

  } else {

    sheet.clearContents();
  }

  return sheet;
}

function microsToCurrency(micros) {

  return Number(micros || 0) / 1000000;
}

function safeDivide(a, b) {

  if (!b || b === 0) {
    return 0;
  }

  return a / b;
}

function pctString(value) {

  return Number(
    (value * 100).toFixed(2)
  ) + '%';
}

/**
 * =========================================================
 * FETCH CONVERSION MAP — CAMPAIGN LEVEL
 * =========================================================
 */

function getConversionMap() {

  const map = {};

  const query = `
    SELECT
      segments.month,
      campaign.name,
      segments.conversion_action_name,
      metrics.all_conversions_by_conversion_date,
      metrics.all_conversions_value_by_conversion_date
    FROM campaign
    WHERE
      segments.date BETWEEN '${DATE_FROM}' AND '${DATE_TO}'
  `;

  const rows = AdsApp.search(query);

  while (rows.hasNext()) {

    const row = rows.next();

    const month =
      String(row.segments.month || '')
      .trim();

    const campaign =
      String(row.campaign.name || '')
      .trim();

    const conversionName =
      String(
        row.segments.conversionActionName || ''
      ).trim();

    const key =
      month + '|' + campaign;

    if (!map[key]) {

      map[key] = {
        conv: 0,
        revenue: 0,
        atc: 0
      };
    }

    const conversions =
      Number(
        row.metrics
          .allConversionsByConversionDate || 0
      );

    const revenue =
      Number(
        row.metrics
          .allConversionsValueByConversionDate || 0
      );

    /**
     * GA4 PURCHASE
     */

    if (
      conversionName.indexOf(
        CONV.PURCHASE
      ) !== -1
    ) {

      map[key].conv += conversions;
      map[key].revenue += revenue;
    }

    /**
     * ATC
     */

    if (
      conversionName.indexOf(
        CONV.ATC
      ) !== -1
    ) {

      map[key].atc += conversions;
    }
  }

  return map;
}

/**
 * =========================================================
 * FETCH CONVERSION MAP — AD GROUP LEVEL
 * =========================================================
 */

function getAdGroupConversionMap() {

  const map = {};

  const query = `
    SELECT
      segments.month,
      campaign.name,
      ad_group.name,
      segments.conversion_action_name,
      metrics.all_conversions_by_conversion_date,
      metrics.all_conversions_value_by_conversion_date
    FROM ad_group
    WHERE
      segments.date BETWEEN '${DATE_FROM}' AND '${DATE_TO}'
  `;

  const rows = AdsApp.search(query);

  while (rows.hasNext()) {

    const row = rows.next();

    const month =
      String(row.segments.month || '')
      .trim();

    const campaign =
      String(row.campaign.name || '')
      .trim();

    const adGroup =
      String(row.adGroup.name || '')
      .trim();

    const conversionName =
      String(
        row.segments.conversionActionName || ''
      ).trim();

    const key =
      month + '|' + campaign + '|' + adGroup;

    if (!map[key]) {

      map[key] = {
        conv: 0,
        revenue: 0,
        atc: 0
      };
    }

    const conversions =
      Number(
        row.metrics
          .allConversionsByConversionDate || 0
      );

    const revenue =
      Number(
        row.metrics
          .allConversionsValueByConversionDate || 0
      );

    if (
      conversionName.indexOf(
        CONV.PURCHASE
      ) !== -1
    ) {
      map[key].conv += conversions;
      map[key].revenue += revenue;
    }

    if (
      conversionName.indexOf(
        CONV.ATC
      ) !== -1
    ) {
      map[key].atc += conversions;
    }
  }

  return map;
}

/**
 * =========================================================
 * CAMPAIGN REPORT
 * =========================================================
 */

function exportCampaignReport(
  spreadsheet
) {

  const sheet =
    getOrCreateSheet(
      spreadsheet,
      'Campaign Report'
    );

  const headers = [

    'Month',
    'Campaign',

    'Spend',
    'Clicks',
    'Impressions',
    'CTR',
    'Avg CPC',

    'Conv (by conv. time)',
    'Revenue (by conv. time)',

    'ROAS (by conv. time)',
    'CPA (by conv. time)',

    'ATC'
  ];

  sheet.appendRow(headers);

  const conversionMap =
    getConversionMap();

  const query = `
    SELECT
      segments.month,
      campaign.name,
      metrics.impressions,
      metrics.clicks,
      metrics.ctr,
      metrics.average_cpc,
      metrics.cost_micros
    FROM campaign
    WHERE
      segments.date BETWEEN '${DATE_FROM}' AND '${DATE_TO}'
    ORDER BY metrics.cost_micros DESC
  `;

  const rows = AdsApp.search(query);

  while (rows.hasNext()) {

    const row = rows.next();

    const month =
      row.segments.month;

    const campaign =
      row.campaign.name;

    const key =
      month + '|' + campaign;

    const convData =
      conversionMap[key] || {};

    const spend =
      microsToCurrency(
        row.metrics.costMicros
      );

    const conv =
      Number(convData.conv || 0);

    const revenue =
      Number(convData.revenue || 0);

    const roas =
      safeDivide(revenue, spend);

    const cpa =
      safeDivide(spend, conv);

    sheet.appendRow([

      month,
      campaign,

      Number(spend.toFixed(2)),
      row.metrics.clicks,
      row.metrics.impressions,

      Number(
        (row.metrics.ctr * 100).toFixed(2)
      ) + '%',

      Number(
        microsToCurrency(
          row.metrics.averageCpc
        ).toFixed(2)
      ),

      Number(conv.toFixed(2)),
      Math.round(revenue),

      Number(roas.toFixed(2)),
      Number(cpa.toFixed(2)),

      Math.round(convData.atc || 0)
    ]);
  }

  Logger.log('Campaign Report Exported');
}

/**
 * =========================================================
 * AD GROUP REPORT
 * =========================================================
 */

function exportAdGroupReport(
  spreadsheet
) {

  const sheet =
    getOrCreateSheet(
      spreadsheet,
      'Ad Group Report'
    );

  const headers = [
    'Month',
    'Campaign',
    'Ad Group',
    'Spend',
    'Clicks',
    'Impressions',
    'CTR',
    'Avg CPC',
    'Conv (by conv. time)',
    'Revenue (by conv. time)',
    'ROAS (by conv. time)',
    'CPA (by conv. time)',
    'ATC'
  ];

  sheet.appendRow(headers);

  const conversionMap =
    getAdGroupConversionMap();

  const query = `
    SELECT
      segments.month,
      campaign.name,
      ad_group.name,
      metrics.impressions,
      metrics.clicks,
      metrics.ctr,
      metrics.average_cpc,
      metrics.cost_micros
    FROM ad_group
    WHERE
      segments.date BETWEEN '${DATE_FROM}' AND '${DATE_TO}'
    ORDER BY metrics.cost_micros DESC
  `;

  const rows = AdsApp.search(query);

  while (rows.hasNext()) {

    const row = rows.next();

    const month = row.segments.month;
    const campaign = row.campaign.name;
    const adGroup = row.adGroup.name;

    const key =
      month + '|' + campaign + '|' + adGroup;

    const convData =
      conversionMap[key] || {};

    const spend =
      microsToCurrency(row.metrics.costMicros);

    const conv =
      Number(convData.conv || 0);

    const revenue =
      Number(convData.revenue || 0);

    const roas =
      safeDivide(revenue, spend);

    const cpa =
      safeDivide(spend, conv);

    sheet.appendRow([
      month,
      campaign,
      adGroup,
      Number(spend.toFixed(2)),
      row.metrics.clicks,
      row.metrics.impressions,
      Number(
        (row.metrics.ctr * 100).toFixed(2)
      ) + '%',
      Number(
        microsToCurrency(
          row.metrics.averageCpc
        ).toFixed(2)
      ),
      Number(conv.toFixed(2)),
      Math.round(revenue),
      Number(roas.toFixed(2)),
      Number(cpa.toFixed(2)),
      Math.round(convData.atc || 0)
    ]);
  }

  Logger.log('Ad Group Report Exported');
}

/**
 * =========================================================
 * KEYWORD REPORT
 * =========================================================
 *
 * Note: keyword_view supports regular all_conversions
 * but not the new tracker split via segmentation
 * (would explode row count). Uses click-date attribution.
 */

function exportKeywordReport(
  spreadsheet
) {

  const sheet =
    getOrCreateSheet(
      spreadsheet,
      'Keyword Report'
    );

  const query = `
    SELECT
      segments.month,
      campaign.name,
      ad_group.name,
      ad_group_criterion.keyword.text,
      ad_group_criterion.keyword.match_type,
      metrics.impressions,
      metrics.clicks,
      metrics.ctr,
      metrics.average_cpc,
      metrics.cost_micros,
      metrics.all_conversions,
      metrics.all_conversions_value
    FROM keyword_view
    WHERE
      segments.date BETWEEN '${DATE_FROM}' AND '${DATE_TO}'
      AND ad_group_criterion.status = 'ENABLED'
    ORDER BY metrics.cost_micros DESC
  `;

  const report = AdsApp.report(query);
  report.exportToSheet(sheet);

  Logger.log('Keyword Report Exported');
}

/**
 * =========================================================
 * SEARCH TERMS REPORT  (click-date attribution)
 * =========================================================
 */

function exportSearchTermsReport(
  spreadsheet
) {

  const sheet =
    getOrCreateSheet(
      spreadsheet,
      'Search Terms'
    );

  const query = `
    SELECT
      segments.month,
      campaign.name,
      ad_group.name,
      search_term_view.search_term,
      metrics.impressions,
      metrics.clicks,
      metrics.ctr,
      metrics.average_cpc,
      metrics.cost_micros,
      metrics.all_conversions,
      metrics.all_conversions_value
    FROM search_term_view
    WHERE
      segments.date BETWEEN '${DATE_FROM}' AND '${DATE_TO}'
    ORDER BY metrics.clicks DESC
  `;

  const report = AdsApp.report(query);
  report.exportToSheet(sheet);

  Logger.log('Search Terms Exported');
}

/**
 * =========================================================
 * PRODUCT PERFORMANCE REPORT  (click-date attribution)
 * =========================================================
 */

function exportProductReport(
  spreadsheet
) {

  const sheet =
    getOrCreateSheet(
      spreadsheet,
      'Product Performance'
    );

  const query = `
    SELECT
      segments.month,
      campaign.name,
      segments.product_title,
      segments.product_item_id,
      metrics.impressions,
      metrics.clicks,
      metrics.cost_micros,
      metrics.all_conversions,
      metrics.all_conversions_value
    FROM shopping_performance_view
    WHERE
      segments.date BETWEEN '${DATE_FROM}' AND '${DATE_TO}'
    ORDER BY metrics.cost_micros DESC
  `;

  const report = AdsApp.report(query);
  report.exportToSheet(sheet);

  Logger.log('Product Performance Exported');
}

/**
 * =========================================================
 * DEVICE REPORT  (click-date attribution)
 * =========================================================
 */

function exportDeviceReport(
  spreadsheet
) {

  const sheet =
    getOrCreateSheet(
      spreadsheet,
      'Device Report'
    );

  const query = `
    SELECT
      segments.month,
      segments.device,
      campaign.name,
      metrics.impressions,
      metrics.clicks,
      metrics.ctr,
      metrics.average_cpc,
      metrics.cost_micros,
      metrics.all_conversions,
      metrics.all_conversions_value
    FROM campaign
    WHERE
      segments.date BETWEEN '${DATE_FROM}' AND '${DATE_TO}'
    ORDER BY metrics.cost_micros DESC
  `;

  const report = AdsApp.report(query);
  report.exportToSheet(sheet);

  Logger.log('Device Report Exported');
}

/**
 * =========================================================
 * GEO REPORT  (click-date attribution)
 * =========================================================
 */

function exportGeoReport(
  spreadsheet
) {

  const sheet =
    getOrCreateSheet(
      spreadsheet,
      'Geo Report'
    );

  const query = `
    SELECT
      campaign.name,
      geographic_view.country_criterion_id,
      metrics.impressions,
      metrics.clicks,
      metrics.cost_micros,
      metrics.all_conversions,
      metrics.all_conversions_value
    FROM geographic_view
    WHERE
      segments.date BETWEEN '${DATE_FROM}' AND '${DATE_TO}'
    ORDER BY metrics.cost_micros DESC
  `;

  const report = AdsApp.report(query);
  report.exportToSheet(sheet);

  Logger.log('Geo Report Exported');
}

/**
 * =========================================================
 * DAY OF WEEK PERFORMANCE
 * =========================================================
 *
 * Aggregates spend / clicks / conv by day of week
 * across the date range. Useful for finding which
 * days underperform / overperform → ad scheduling.
 *
 * Uses click-date all_conversions because this
 * report blends all conversion actions for the
 * scheduling view; not filtered to GA4 Purchase.
 */

function exportDayOfWeekReport(
  spreadsheet
) {

  const sheet =
    getOrCreateSheet(
      spreadsheet,
      'Day of Week'
    );

  const headers = [
    'Day of Week',
    'Campaign',
    'Spend',
    'Clicks',
    'Impressions',
    'CTR',
    'Avg CPC',
    'Conv',
    'Revenue',
    'ROAS',
    'Cost / Conv'
  ];

  sheet.appendRow(headers);

  const query = `
    SELECT
      segments.day_of_week,
      campaign.name,
      metrics.impressions,
      metrics.clicks,
      metrics.ctr,
      metrics.average_cpc,
      metrics.cost_micros,
      metrics.all_conversions,
      metrics.all_conversions_value
    FROM campaign
    WHERE
      segments.date BETWEEN '${DATE_FROM}' AND '${DATE_TO}'
    ORDER BY segments.day_of_week, metrics.cost_micros DESC
  `;

  const rows = AdsApp.search(query);

  while (rows.hasNext()) {

    const row = rows.next();

    const spend =
      microsToCurrency(row.metrics.costMicros);

    const conv =
      Number(row.metrics.allConversions || 0);

    const revenue =
      Number(
        row.metrics.allConversionsValue || 0
      );

    const roas = safeDivide(revenue, spend);
    const costPerConv = safeDivide(spend, conv);

    sheet.appendRow([
      row.segments.dayOfWeek,
      row.campaign.name,
      Number(spend.toFixed(2)),
      row.metrics.clicks,
      row.metrics.impressions,
      Number(
        (row.metrics.ctr * 100).toFixed(2)
      ) + '%',
      Number(
        microsToCurrency(
          row.metrics.averageCpc
        ).toFixed(2)
      ),
      Number(conv.toFixed(2)),
      Number(revenue.toFixed(2)),
      Number(roas.toFixed(2)),
      Number(costPerConv.toFixed(2))
    ]);
  }

  Logger.log('Day of Week Report Exported');
}

/**
 * =========================================================
 * FUNNEL ANALYSIS
 * =========================================================
 *
 * Derived insights — no extra API calls.
 * Reuses the campaign-level conversion map.
 *
 * Per campaign × month, computes:
 *   - ATC count
 *   - Purchase count (GA4 Hunnit web purchase)
 *   - ATC → Purchase rate
 *   - Spend / ATC, Spend / Purchase
 *
 * Helps spot leaks: low ATC = creative/targeting;
 * ATC fine but Purchase drops = checkout/payment
 * friction.
 */

function exportFunnelAnalysis(
  spreadsheet
) {

  const sheet =
    getOrCreateSheet(
      spreadsheet,
      'Funnel Analysis'
    );

  const headers = [
    'Month',
    'Campaign',
    'Spend',
    'ATC',
    'Purchases',
    'ATC → Purchase %',
    'Cost / ATC',
    'Cost / Purchase'
  ];

  sheet.appendRow(headers);

  const conversionMap = getConversionMap();

  // Pull spend separately to pair with conv map
  const spendByKey = {};

  const query = `
    SELECT
      segments.month,
      campaign.name,
      metrics.cost_micros
    FROM campaign
    WHERE
      segments.date BETWEEN '${DATE_FROM}' AND '${DATE_TO}'
  `;

  const rows = AdsApp.search(query);

  while (rows.hasNext()) {
    const row = rows.next();
    const key =
      row.segments.month + '|' + row.campaign.name;
    spendByKey[key] =
      (spendByKey[key] || 0) +
      microsToCurrency(row.metrics.costMicros);
  }

  // Build funnel rows
  const funnelRows = [];

  Object.keys(conversionMap).forEach(
    function (key) {

      const data = conversionMap[key];
      const parts = key.split('|');
      const month = parts[0];
      const campaign = parts[1];

      const spend = spendByKey[key] || 0;

      const atc = Number(data.atc || 0);
      const purchases = Number(data.conv || 0);

      const atcToPurchase =
        safeDivide(purchases, atc);

      const costPerAtc =
        safeDivide(spend, atc);

      const costPerPurchase =
        safeDivide(spend, purchases);

      funnelRows.push([
        month,
        campaign,
        Number(spend.toFixed(2)),
        Math.round(atc),
        Math.round(purchases),
        pctString(atcToPurchase),
        Number(costPerAtc.toFixed(2)),
        Number(costPerPurchase.toFixed(2))
      ]);
    }
  );

  // Sort by spend desc for readability
  funnelRows.sort(function (a, b) {
    return b[2] - a[2];
  });

  funnelRows.forEach(function (r) {
    sheet.appendRow(r);
  });

  Logger.log('Funnel Analysis Exported');
}
