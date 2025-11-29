/**
 * ============================================================================
 * PREDICTIVE ANALYTICS
 * ============================================================================
 *
 * Forecasts trends and provides actionable insights
 * Features:
 * - Grievance volume forecasting
 * - Issue type trend analysis
 * - Seasonal pattern detection
 * - Steward workload predictions
 * - Risk identification
 * - Anomaly detection
 * - Actionable recommendations
 */

/**
 * Analyzes trends and generates predictions
 * @returns {Object} Analytics results
 */
function performPredictiveAnalysis() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const grievanceSheet = ss.getSheetByName(SHEETS.GRIEVANCE_LOG);
  const lastRow = grievanceSheet.getLastRow();

  if (lastRow < 2) {
    return {
      error: 'Insufficient data for analysis'
    };
  }

  const data = grievanceSheet.getRange(2, 1, lastRow - 1, 28).getValues();

  const analytics = {
    volumeTrend: analyzeVolumeTrend(data),
    issueTypeTrends: analyzeIssueTypeTrends(data),
    seasonalPatterns: detectSeasonalPatterns(data),
    resolutionTimeTrend: analyzeResolutionTimeTrend(data),
    stewardWorkloadForecast: forecastStewardWorkload(data),
    riskFactors: identifyRiskFactors(data),
    anomalies: detectAnomalies(data),
    recommendations: generateRecommendations(data)
  };

  return analytics;
}

/**
 * Analyzes grievance volume trend
 * @param {Array} data - Grievance data
 * @returns {Object} Volume trend analysis
 */
function analyzeVolumeTrend(data) {
  // Group by month
  const monthlyVolumes = {};

  data.forEach(row => {
    const filedDate = row[6]; // Column G: Filed Date

    if (!filedDate) return;

    const monthKey = `${filedDate.getFullYear()}-${String(filedDate.getMonth() + 1).padStart(2, '0')}`;

    monthlyVolumes[monthKey] = (monthlyVolumes[monthKey] || 0) + 1;
  });

  // Convert to array and sort
  const months = Object.keys(monthlyVolumes).sort();
  const volumes = months.map(m => monthlyVolumes[m]);

  // Calculate trend (simple linear regression)
  const trend = calculateLinearTrend(volumes);

  // Forecast next 3 months
  const forecast = [];
  for (let i = 1; i <= 3; i++) {
    forecast.push(Math.round(trend.slope * (volumes.length + i) + trend.intercept));
  }

  return {
    currentMonthVolume: volumes[volumes.length - 1] || 0,
    previousMonthVolume: volumes[volumes.length - 2] || 0,
    trend: trend.slope > 0 ? 'increasing' : trend.slope < 0 ? 'decreasing' : 'stable',
    trendPercentage: Math.abs(trend.slope),
    forecast: forecast,
    monthlyVolumes: monthlyVolumes
  };
}

/**
 * Analyzes issue type trends
 * @param {Array} data - Grievance data
 * @returns {Object} Issue type trends
 */
function analyzeIssueTypeTrends(data) {
  const issueTypesByMonth = {};

  data.forEach(row => {
    const filedDate = row[6];
    const issueType = row[5];

    if (!filedDate || !issueType) return;

    const monthKey = `${filedDate.getFullYear()}-${String(filedDate.getMonth() + 1).padStart(2, '0')}`;

    if (!issueTypesByMonth[monthKey]) {
      issueTypesByMonth[monthKey] = {};
    }

    issueTypesByMonth[monthKey][issueType] = (issueTypesByMonth[monthKey][issueType] || 0) + 1;
  });

  // Find trending issue types (increasing over time)
  const trendingIssues = {};
  const months = Object.keys(issueTypesByMonth).sort();

  if (months.length >= 3) {
    const recentMonths = months.slice(-3);
    const olderMonths = months.slice(0, -3);

    const recentCounts = {};
    const olderCounts = {};

    recentMonths.forEach(month => {
      Object.entries(issueTypesByMonth[month]).forEach(([type, count]) => {
        recentCounts[type] = (recentCounts[type] || 0) + count;
      });
    });

    olderMonths.forEach(month => {
      Object.entries(issueTypesByMonth[month]).forEach(([type, count]) => {
        olderCounts[type] = (olderCounts[type] || 0) + count;
      });
    });

    Object.keys(recentCounts).forEach(type => {
      const recentAvg = recentCounts[type] / recentMonths.length;
      const olderAvg = olderCounts[type] ? olderCounts[type] / olderMonths.length : 0;

      if (recentAvg > olderAvg * 1.2) {
        trendingIssues[type] = {
          recentAverage: recentAvg.toFixed(1),
          olderAverage: olderAvg.toFixed(1),
          percentIncrease: olderAvg > 0 ? (((recentAvg - olderAvg) / olderAvg) * 100).toFixed(1) : 100
        };
      }
    });
  }

  return {
    trendingUp: trendingIssues,
    monthlyBreakdown: issueTypesByMonth
  };
}

/**
 * Detects seasonal patterns
 * @param {Array} data - Grievance data
 * @returns {Object} Seasonal patterns
 */
function detectSeasonalPatterns(data) {
  const quarterlyVolumes = { Q1: 0, Q2: 0, Q3: 0, Q4: 0 };

  data.forEach(row => {
    const filedDate = row[6];

    if (!filedDate) return;

    const quarter = Math.floor(filedDate.getMonth() / 3) + 1;
    quarterlyVolumes[`Q${quarter}`]++;
  });

  // Find peak quarter
  let peakQuarter = 'Q1';
  let peakVolume = 0;

  Object.entries(quarterlyVolumes).forEach(([quarter, volume]) => {
    if (volume > peakVolume) {
      peakQuarter = quarter;
      peakVolume = volume;
    }
  });

  return {
    quarterlyVolumes: quarterlyVolumes,
    peakQuarter: peakQuarter,
    peakVolume: peakVolume,
    hasSeasonality: Math.max(...Object.values(quarterlyVolumes)) > Math.min(...Object.values(quarterlyVolumes)) * 1.5
  };
}

/**
 * Analyzes resolution time trend
 * @param {Array} data - Grievance data
 * @returns {Object} Resolution time analysis
 */
function analyzeResolutionTimeTrend(data) {
  const resolutionTimes = [];

  data.forEach(row => {
    const filedDate = row[6];
    const closedDate = row[18];

    if (filedDate && closedDate && closedDate > filedDate) {
      const days = Math.floor((closedDate - filedDate) / (1000 * 60 * 60 * 24));
      resolutionTimes.push(days);
    }
  });

  if (resolutionTimes.length === 0) {
    return {
      average: 0,
      median: 0,
      trend: 'insufficient_data'
    };
  }

  // Sort for median
  resolutionTimes.sort((a, b) => a - b);

  const average = resolutionTimes.reduce((sum, val) => sum + val, 0) / resolutionTimes.length;
  const median = resolutionTimes[Math.floor(resolutionTimes.length / 2)];

  // Recent vs older comparison
  const recent = resolutionTimes.slice(-10);
  const older = resolutionTimes.slice(0, -10);

  const recentAvg = recent.reduce((sum, val) => sum + val, 0) / recent.length;
  const olderAvg = older.length > 0 ? older.reduce((sum, val) => sum + val, 0) / older.length : recentAvg;

  return {
    average: Math.round(average),
    median: median,
    trend: recentAvg > olderAvg ? 'increasing' : recentAvg < olderAvg ? 'decreasing' : 'stable',
    recentAverage: Math.round(recentAvg),
    olderAverage: Math.round(olderAvg)
  };
}

/**
 * Forecasts steward workload
 * @param {Array} data - Grievance data
 * @returns {Object} Workload forecast
 */
function forecastStewardWorkload(data) {
  const stewardCases = {};

  data.forEach(row => {
    const steward = row[13];
    const status = row[4];

    if (steward && status === 'Open') {
      stewardCases[steward] = (stewardCases[steward] || 0) + 1;
    }
  });

  // Identify overloaded stewards (>15 cases)
  const overloaded = [];
  const underutilized = [];

  Object.entries(stewardCases).forEach(([steward, count]) => {
    if (count > 15) {
      overloaded.push({ steward, caseload: count });
    } else if (count < 3) {
      underutilized.push({ steward, caseload: count });
    }
  });

  return {
    overloaded: overloaded,
    underutilized: underutilized,
    balanceRecommendation: overloaded.length > 0 && underutilized.length > 0
      ? 'Redistribute cases from overloaded to underutilized stewards'
      : 'Workload is relatively balanced'
  };
}

/**
 * Identifies risk factors
 * @param {Array} data - Grievance data
 * @returns {Array} Risk factors
 */
function identifyRiskFactors(data) {
  const risks = [];

  // Check for overdue cases
  let overdueCount = 0;
  data.forEach(row => {
    const daysToDeadline = row[20];
    if (daysToDeadline < 0) overdueCount++;
  });

  if (overdueCount > 0) {
    risks.push({
      type: 'Overdue Cases',
      severity: overdueCount > 10 ? 'HIGH' : overdueCount > 5 ? 'MEDIUM' : 'LOW',
      description: `${overdueCount} grievance(s) are overdue`,
      action: 'Review and prioritize overdue cases immediately'
    });
  }

  // Check for cases with no steward assigned
  let unassignedCount = 0;
  data.forEach(row => {
    const steward = row[13];
    const status = row[4];
    if (!steward && status === 'Open') unassignedCount++;
  });

  if (unassignedCount > 0) {
    risks.push({
      type: 'Unassigned Cases',
      severity: unassignedCount > 5 ? 'HIGH' : 'MEDIUM',
      description: `${unassignedCount} open grievance(s) have no assigned steward`,
      action: 'Use auto-assignment to assign stewards'
    });
  }

  // Check for high volume increase
  const volumeTrend = analyzeVolumeTrend(data);
  if (volumeTrend.trend === 'increasing' && volumeTrend.trendPercentage > 20) {
    risks.push({
      type: 'Volume Surge',
      severity: 'MEDIUM',
      description: `Grievance volume trending upward by ${volumeTrend.trendPercentage.toFixed(1)}%`,
      action: 'Prepare for increased caseload, consider additional steward training'
    });
  }

  return risks;
}

/**
 * Detects anomalies in data
 * @param {Array} data - Grievance data
 * @returns {Array} Anomalies detected
 */
function detectAnomalies(data) {
  const anomalies = [];

  // Check for unusual spike in last month
  const monthlyVolumes = {};
  data.forEach(row => {
    const filedDate = row[6];
    if (!filedDate) return;

    const monthKey = `${filedDate.getFullYear()}-${String(filedDate.getMonth() + 1).padStart(2, '0')}`;
    monthlyVolumes[monthKey] = (monthlyVolumes[monthKey] || 0) + 1;
  });

  const volumes = Object.values(monthlyVolumes);
  if (volumes.length >= 3) {
    const average = volumes.slice(0, -1).reduce((sum, val) => sum + val, 0) / (volumes.length - 1);
    const latest = volumes[volumes.length - 1];

    if (latest > average * 1.5) {
      anomalies.push({
        type: 'Volume Spike',
        description: `Current month has ${latest} cases vs average of ${average.toFixed(1)}`,
        impact: 'May indicate systemic issue or policy change'
      });
    }
  }

  // Check for concentration of cases in one location
  const locationCounts = {};
  data.forEach(row => {
    const location = row[9];
    if (location) {
      locationCounts[location] = (locationCounts[location] || 0) + 1;
    }
  });

  const totalCases = Object.values(locationCounts).reduce((sum, val) => sum + val, 0);
  Object.entries(locationCounts).forEach(([location, count]) => {
    if (count > totalCases * 0.4) {
      anomalies.push({
        type: 'Location Concentration',
        description: `${location} has ${count} cases (${((count / totalCases) * 100).toFixed(1)}% of all cases)`,
        impact: 'May indicate location-specific issues requiring attention'
      });
    }
  });

  return anomalies;
}

/**
 * Generates actionable recommendations
 * @param {Array} data - Grievance data
 * @returns {Array} Recommendations
 */
function generateRecommendations(data) {
  const recommendations = [];

  // Based on volume trend
  const volumeTrend = analyzeVolumeTrend(data);
  if (volumeTrend.trend === 'increasing') {
    recommendations.push({
      priority: 'HIGH',
      category: 'Capacity Planning',
      recommendation: 'Grievance volume is increasing. Consider recruiting and training additional stewards.',
      expectedImpact: 'Prevent case backlog and maintain service quality'
    });
  }

  // Based on resolution time
  const resolutionTrend = analyzeResolutionTimeTrend(data);
  if (resolutionTrend.average > 30) {
    recommendations.push({
      priority: 'MEDIUM',
      category: 'Process Improvement',
      recommendation: `Average resolution time is ${resolutionTrend.average} days. Implement workflow automation and deadline reminders.`,
      expectedImpact: 'Reduce resolution time by 20-30%'
    });
  }

  // Based on workload
  const workloadForecast = forecastStewardWorkload(data);
  if (workloadForecast.overloaded.length > 0) {
    recommendations.push({
      priority: 'HIGH',
      category: 'Workload Management',
      recommendation: `${workloadForecast.overloaded.length} steward(s) are overloaded. ${workloadForecast.balanceRecommendation}`,
      expectedImpact: 'Prevent steward burnout and improve case handling'
    });
  }

  // Based on trending issues
  const issueTrends = analyzeIssueTypeTrends(data);
  const trendingIssueCount = Object.keys(issueTrends.trendingUp).length;
  if (trendingIssueCount > 0) {
    const topTrending = Object.keys(issueTrends.trendingUp)[0];
    recommendations.push({
      priority: 'MEDIUM',
      category: 'Root Cause Analysis',
      recommendation: `${topTrending} cases are trending upward. Conduct root cause analysis and consider preventive measures.`,
      expectedImpact: 'Reduce future grievances through proactive intervention'
    });
  }

  return recommendations;
}

/**
 * Calculates simple linear trend
 * @param {Array} values - Data points
 * @returns {Object} Trend line parameters
 */
function calculateLinearTrend(values) {
  const n = values.length;
  if (n < 2) return { slope: 0, intercept: values[0] || 0 };

  let sumX = 0, sumY = 0, sumXY = 0, sumX2 = 0;

  for (let i = 0; i < n; i++) {
    sumX += i;
    sumY += values[i];
    sumXY += i * values[i];
    sumX2 += i * i;
  }

  const slope = (n * sumXY - sumX * sumY) / (n * sumX2 - sumX * sumX);
  const intercept = (sumY - slope * sumX) / n;

  return { slope, intercept };
}

/**
 * Shows predictive analytics dashboard
 */
function showPredictiveAnalyticsDashboard() {
  const ui = SpreadsheetApp.getUi();

  SpreadsheetApp.getActiveSpreadsheet().toast(
    '🔮 Analyzing trends...',
    'Predictive Analytics',
    -1
  );

  try {
    const analytics = performPredictiveAnalysis();

    if (analytics.error) {
      ui.alert('⚠️ ' + analytics.error);
      return;
    }

    const html = createAnalyticsDashboardHTML(analytics);
    const htmlOutput = HtmlService.createHtmlOutput(html)
      .setWidth(900)
      .setHeight(700);

    ui.showModalDialog(htmlOutput, '🔮 Predictive Analytics Dashboard');

  } catch (error) {
    ui.alert('❌ Error', error.message, ui.ButtonSet.OK);
  }
}

/**
 * Creates HTML for analytics dashboard
 * @param {Object} analytics - Analytics results
 * @returns {string} HTML content
 */
function createAnalyticsDashboardHTML(analytics) {
  const risksHTML = analytics.riskFactors.length > 0
    ? analytics.riskFactors.map(risk => `
        <div class="risk-item severity-${risk.severity.toLowerCase()}">
          <div class="risk-header">
            <span class="risk-type">${risk.type}</span>
            <span class="risk-severity">${risk.severity}</span>
          </div>
          <div class="risk-desc">${risk.description}</div>
          <div class="risk-action">⚡ ${risk.action}</div>
        </div>
      `).join('')
    : '<p>No significant risks identified. ✅</p>';

  const recommendationsHTML = analytics.recommendations.length > 0
    ? analytics.recommendations.map(rec => `
        <div class="recommendation priority-${rec.priority.toLowerCase()}">
          <div class="rec-header">
            <span class="rec-category">${rec.category}</span>
            <span class="rec-priority">${rec.priority}</span>
          </div>
          <div class="rec-text">${rec.recommendation}</div>
          <div class="rec-impact">📊 ${rec.expectedImpact}</div>
        </div>
      `).join('')
    : '<p>No specific recommendations at this time.</p>';

  const anomaliesHTML = analytics.anomalies.length > 0
    ? analytics.anomalies.map(anomaly => `
        <div class="anomaly-item">
          <div class="anomaly-type">${anomaly.type}</div>
          <div class="anomaly-desc">${anomaly.description}</div>
          <div class="anomaly-impact">${anomaly.impact}</div>
        </div>
      `).join('')
    : '<p>No anomalies detected. ✅</p>';

  return `
<!DOCTYPE html>
<html>
<head>
  <base target="_top">
  <style>
    body { font-family: 'Roboto', Arial, sans-serif; padding: 20px; margin: 0; background: #f5f5f5; }
    .container { background: white; padding: 25px; border-radius: 8px; box-shadow: 0 2px 8px rgba(0,0,0,0.1); max-height: 650px; overflow-y: auto; }
    h2 { color: #1a73e8; margin-top: 0; border-bottom: 3px solid #1a73e8; padding-bottom: 10px; }
    h3 { color: #333; margin-top: 25px; border-bottom: 2px solid #e0e0e0; padding-bottom: 8px; }
    .summary-grid { display: grid; grid-template-columns: repeat(3, 1fr); gap: 15px; margin: 20px 0; }
    .summary-card { background: #f8f9fa; padding: 15px; border-radius: 8px; border-left: 4px solid #1a73e8; }
    .summary-label { font-size: 12px; color: #666; text-transform: uppercase; }
    .summary-value { font-size: 24px; font-weight: bold; color: #333; margin: 5px 0; }
    .summary-trend { font-size: 13px; color: #666; }
    .trend-up { color: #f44336; }
    .trend-down { color: #4caf50; }
    .trend-stable { color: #9e9e9e; }
    .risk-item { background: #fff3e0; padding: 15px; margin: 10px 0; border-radius: 4px; border-left: 4px solid; }
    .risk-item.severity-high { border-left-color: #f44336; background: #ffebee; }
    .risk-item.severity-medium { border-left-color: #ff9800; background: #fff3e0; }
    .risk-item.severity-low { border-left-color: #ffeb3b; background: #fffde7; }
    .risk-header { display: flex; justify-content: space-between; margin-bottom: 8px; }
    .risk-type { font-weight: bold; color: #333; }
    .risk-severity { background: white; padding: 2px 10px; border-radius: 12px; font-size: 12px; font-weight: bold; }
    .risk-desc { margin: 5px 0; color: #555; }
    .risk-action { margin-top: 10px; padding: 10px; background: white; border-radius: 4px; font-size: 13px; }
    .recommendation { background: #e8f5e9; padding: 15px; margin: 10px 0; border-radius: 4px; border-left: 4px solid #4caf50; }
    .recommendation.priority-high { border-left-color: #f44336; background: #ffebee; }
    .recommendation.priority-medium { border-left-color: #ff9800; background: #fff3e0; }
    .rec-header { display: flex; justify-content: space-between; margin-bottom: 8px; }
    .rec-category { font-weight: bold; color: #333; }
    .rec-priority { background: white; padding: 2px 10px; border-radius: 12px; font-size: 12px; font-weight: bold; }
    .rec-text { margin: 5px 0; color: #555; }
    .rec-impact { margin-top: 10px; padding: 10px; background: white; border-radius: 4px; font-size: 13px; }
    .anomaly-item { background: #e3f2fd; padding: 15px; margin: 10px 0; border-radius: 4px; border-left: 4px solid #2196f3; }
    .anomaly-type { font-weight: bold; color: #1976d2; margin-bottom: 5px; }
    .anomaly-desc { color: #555; margin: 5px 0; }
    .anomaly-impact { margin-top: 10px; padding: 10px; background: white; border-radius: 4px; font-size: 13px; font-style: italic; }
  </style>
</head>
<body>
  <div class="container">
    <h2>🔮 Predictive Analytics Dashboard</h2>

    <div class="summary-grid">
      <div class="summary-card">
        <div class="summary-label">Volume Trend</div>
        <div class="summary-value trend-${analytics.volumeTrend.trend === 'increasing' ? 'up' : analytics.volumeTrend.trend === 'decreasing' ? 'down' : 'stable'}">
          ${analytics.volumeTrend.currentMonthVolume}
        </div>
        <div class="summary-trend">
          ${analytics.volumeTrend.trend === 'increasing' ? '📈 Increasing' : analytics.volumeTrend.trend === 'decreasing' ? '📉 Decreasing' : '➡️ Stable'}
        </div>
      </div>

      <div class="summary-card">
        <div class="summary-label">Avg Resolution Time</div>
        <div class="summary-value">${analytics.resolutionTimeTrend.average} days</div>
        <div class="summary-trend">
          ${analytics.resolutionTimeTrend.trend === 'increasing' ? '⚠️ Increasing' : analytics.resolutionTimeTrend.trend === 'decreasing' ? '✅ Improving' : '➡️ Stable'}
        </div>
      </div>

      <div class="summary-card">
        <div class="summary-label">Identified Risks</div>
        <div class="summary-value">${analytics.riskFactors.length}</div>
        <div class="summary-trend">
          ${analytics.riskFactors.length === 0 ? '✅ No risks' : `⚠️ Needs attention`}
        </div>
      </div>
    </div>

    <h3>🚨 Risk Factors</h3>
    ${risksHTML}

    <h3>💡 Recommendations</h3>
    ${recommendationsHTML}

    <h3>🔍 Anomalies Detected</h3>
    ${anomaliesHTML}

    <h3>📊 Forecast (Next 3 Months)</h3>
    <div style="background: #f8f9fa; padding: 15px; border-radius: 4px; margin: 10px 0;">
      <strong>Projected Volume:</strong><br>
      Month 1: ${analytics.volumeTrend.forecast[0]} cases<br>
      Month 2: ${analytics.volumeTrend.forecast[1]} cases<br>
      Month 3: ${analytics.volumeTrend.forecast[2]} cases
    </div>

    ${analytics.seasonalPatterns.hasSeasonality ? `
    <h3>📅 Seasonal Patterns</h3>
    <div style="background: #f8f9fa; padding: 15px; border-radius: 4px; margin: 10px 0;">
      <strong>Peak Season:</strong> ${analytics.seasonalPatterns.peakQuarter} (${analytics.seasonalPatterns.peakVolume} cases)<br>
      <strong>Quarterly Breakdown:</strong><br>
      Q1: ${analytics.seasonalPatterns.quarterlyVolumes.Q1} |
      Q2: ${analytics.seasonalPatterns.quarterlyVolumes.Q2} |
      Q3: ${analytics.seasonalPatterns.quarterlyVolumes.Q3} |
      Q4: ${analytics.seasonalPatterns.quarterlyVolumes.Q4}
    </div>
    ` : ''}
  </div>
</body>
</html>
  `;
}
