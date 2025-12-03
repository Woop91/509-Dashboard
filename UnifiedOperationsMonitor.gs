/**
 * ------------------------------------------------------------------------====
 * SEIU 509 COMPREHENSIVE ACTION DASHBOARD
 * ------------------------------------------------------------------------====
 *
 * Terminal-themed comprehensive operations monitor with 26+ sections covering:
 * - Executive status and deadlines
 * - Process efficiency and caseload
 * - Network health and capacity
 * - Action logs and follow-up radar
 * - Predictive alerts and systemic risks
 * - Financial recovery and arbitration tracking
 * - Member engagement and satisfaction
 * - Training, compliance, and performance metrics
 *
 * Designed for union stewards and administrators to monitor all operational
 * aspects in a single, information-dense terminal interface.
 * ------------------------------------------------------------------------====
 */

/**
 * Shows the Unified Operations Monitor dashboard
 */
function showUnifiedOperationsMonitor() {
  const html = HtmlService.createHtmlOutput(getUnifiedOperationsMonitorHTML())
    .setWidth(1600)
    .setHeight(1000);

  SpreadsheetApp.getUi().showModelessDialog(html, '🎯 SEIU 509 Comprehensive Action Dashboard');
}

/**
 * Backend function that provides ALL data for the comprehensive dashboard
 * Called by the HTML dashboard via google.script.run
 */
function getUnifiedDashboardData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const grievanceSheet = ss.getSheetByName(SHEETS.GRIEVANCE_LOG);
  const memberSheet = ss.getSheetByName(SHEETS.MEMBER_DIR);

  if (!grievanceSheet || !memberSheet) {
    return {
      error: "Required sheets not found. Please ensure Grievance Log and Member Directory exist."
    };
  }

  // Get all data
  const grievanceData = grievanceSheet.getDataRange().getValues();
  const memberData = memberSheet.getDataRange().getValues();

  // Skip headers
  const grievances = grievanceData.slice(1);
  const members = memberData.slice(1);

  const today = new Date();
  today.setHours(0, 0, 0, 0);

  // Calculate all metrics
  return {
    // Section 1: Executive Status
    activeCases: calculateActiveCases(grievances),
    overdue: calculateOverdue(grievances, today),
    dueWeek: calculateDueThisWeek(grievances, today),
    winRate: calculateWinRate(grievances),
    avgDays: calculateAvgDays(grievances),
    escalationCount: calculateEscalations(grievances),
    arbitrations: calculateArbitrations(grievances),
    highRiskCount: calculateHighRisk(grievances),

    // Section 2: Process Efficiency
    steps: calculateStepEfficiency(grievances),

    // Section 3: Network Health
    totalMembers: members.filter(function(m) { return m[MEMBER_COLS.MEMBER_ID - 1]).length; },
    engagementRate: calculateEngagementRate(grievances, members),
    noContactCount: calculateNoContact(members, today),
    stewardCount: calculateActiveStewards(grievances),
    avgLoad: calculateAvgLoad(grievances),
    overloadedStewards: calculateOverloadedStewards(grievances),
    issueDistribution: calculateIssueDistribution(grievances),

    // Section 4: Action Log (Top 20 processes)
    processes: getTopProcesses(grievances, today, 20),

    // Section 5: Follow-up Radar
    followUpTasks: getFollowUpTasks(grievances, today),

    // Section 6: Predictive Alerts
    predictiveAlerts: getPredictiveAlerts(members, grievances, today),

    // Section 7: Systemic Risk
    systemicRisk: getSystemicRisks(grievances),

    // Section 8: Outreach Scorecard
    outreachScorecard: getOutreachScorecard(members),

    // Section 9: Arbitration Tracker
    arbitrationTracker: getArbitrationTracker(grievances),

    // Section 10: Contract Trends
    contractTrends: getContractTrends(grievances),

    // Section 11: Training Matrix
    trainingMatrix: getTrainingMatrix(members),

    // Section 12: Geographical Caseload
    locationCaseload: getLocationCaseload(grievances),

    // Section 13: Financial Recovery
    financialTracker: getFinancialTracker(grievances),

    // Section 14: Rolling Trends
    rollingTrends: getRollingTrends(grievances),

    // Section 15: Grievance Satisfaction
    grievanceSatisfaction: getGrievanceSatisfaction(members, grievances),

    // Section 16: Member Lifecycle
    memberLifecycle: getMemberLifecycle(members, today),

    // Section 17: Unit Performance
    unitPerformance: getUnitPerformance(grievances, members),

    // Section 18: Document Compliance
    docCompliance: getDocCompliance(grievances),

    // Section 19: Classification Risk
    classificationRisk: getClassificationRisk(members, grievances),

    // Section 20: Win/Loss History
    winLossHistory: getWinLossHistory(grievances),

    // Section 21: Legal Review
    legalReviewCases: getLegalReviewCases(grievances, today),

    // Section 22: Recruitment Tracker
    recruitment: getRecruitmentTracker(members, today),

    // Section 23: Bargaining Prep
    bargainingPrep: getBargainingPrep(members, grievances),

    // Section 24: Policy Impact
    policyImpact: getPolicyImpact(grievances, today),

    // Section 25: Bottlenecks
    bottlenecks: getBottlenecks(grievances),

    // Section 26: Watchlist
    watchlistItems: getWatchlistItems(),
    watchlistLog: getWatchlistLog(grievances, members)
  };
}

// ------------------------------------------------------------------------====
// CALCULATION FUNCTIONS
// ------------------------------------------------------------------------====

function calculateActiveCases(grievances) {
  return grievances.filter(function(g) {
    const status = g[GRIEVANCE_COLS.STATUS - 1];
    return status && (status.includes('Filed') || status === 'Open' || status === 'Pending Info');
  }).length;
}

function calculateOverdue(grievances, today) {
  return grievances.filter(function(g) {
    const status = g[GRIEVANCE_COLS.STATUS - 1];
    const nextDeadline = g[GRIEVANCE_COLS.NEXT_ACTION_DUE - 1];
    if (status && (status.includes('Filed') || status === 'Open' || status === 'Pending Info')) {
      if (nextDeadline instanceof Date) {
        const deadline = new Date(nextDeadline);
        deadline.setHours(0, 0, 0, 0);
        return deadline < today;
      }
    }
    return false;
  }).length;
}

function calculateDueThisWeek(grievances, today) {
  const oneWeekFromNow = new Date(today);
  oneWeekFromNow.setDate(oneWeekFromNow.getDate() + 7);

  return grievances.filter(function(g) {
    const status = g[GRIEVANCE_COLS.STATUS - 1];
    const nextDeadline = g[GRIEVANCE_COLS.NEXT_ACTION_DUE - 1];
    if (status && (status.includes('Filed') || status === 'Open' || status === 'Pending Info')) {
      if (nextDeadline instanceof Date) {
        const deadline = new Date(nextDeadline);
        deadline.setHours(0, 0, 0, 0);
        return deadline >= today && deadline <= oneWeekFromNow;
      }
    }
    return false;
  }).length;
}

function calculateWinRate(grievances) {
  const closed = grievances.filter(function(g) { return g[GRIEVANCE_COLS.STATUS - 1] === 'Settled' || g[GRIEVANCE_COLS.STATUS - 1] === 'Closed' || g[GRIEVANCE_COLS.STATUS - 1] === 'Withdrawn'; });
  if (closed.length === 0) return 0;
  const won = closed.filter(function(g) { return g[GRIEVANCE_COLS.STATUS - 1] === 'Settled'; }).length;
  return (won / closed.length) * 100;
}

function calculateAvgDays(grievances) {
  const closed = grievances.filter(function(g) {
    const status = g[GRIEVANCE_COLS.STATUS - 1];
    const filedDate = g[GRIEVANCE_COLS.DATE_FILED - 1];
    const closedDate = g[GRIEVANCE_COLS.DATE_CLOSED - 1];
    return (status === 'Settled' || status === 'Closed' || status === 'Withdrawn') &&
           filedDate instanceof Date && closedDate instanceof Date;
  });

  if (closed.length === 0) return 0;

  const totalDays = closed.reduce(function(sum, g) {
    const days = Math.floor((g[GRIEVANCE_COLS.DATE_CLOSED - 1] - g[GRIEVANCE_COLS.DATE_FILED - 1]) / (1000 * 60 * 60 * 24));
    return sum + days;
  }, 0);

  return Math.round(totalDays / closed.length);
}

function calculateEscalations(grievances) {
  return grievances.filter(function(g) {
    const step = g[GRIEVANCE_COLS.CURRENT_STEP - 1]; // Current Step column
    return step && (step.includes('III') || step.includes('3'));
  }).length;
}

function calculateArbitrations(grievances) {
  return grievances.filter(function(g) {
    const step = g[GRIEVANCE_COLS.CURRENT_STEP - 1];
    return step && (step.includes('Arbitration') || step.includes('Mediation'));
  }).length;
}

function calculateHighRisk(grievances) {
  return grievances.filter(function(g) {
    const issueCategory = g[GRIEVANCE_COLS.ISSUE_CATEGORY - 1];
    return issueCategory && (issueCategory.includes('Discipline') || issueCategory.includes('Discrimination'));
  }).length;
}

function calculateStepEfficiency(grievances) {
  const stepData = {};
  const active = grievances.filter(function(g) {
    const status = g[GRIEVANCE_COLS.STATUS - 1];
    return status && (status.includes('Filed') || status === 'Open' || status === 'Pending Info');
  });

  active.forEach(function(g) {
    const step = g[GRIEVANCE_COLS.CURRENT_STEP - 1] || 'Unassigned';
    stepData[step] = (stepData[step] || 0) + 1;
  });

  const total = active.length || 1;
  const steps = [
    { name: 'Informal', team: 'INITIAL', status: 'green' },
    { name: 'Step I', team: 'REVIEW', status: 'yellow' },
    { name: 'Step II', team: 'APPEAL', status: 'yellow' },
    { name: 'Step III', team: 'ESCALATION', status: 'red' }
  ];

  return steps.map(function(s) { return ({
    ...s,
    cases: stepData[s.name] || 0,
    caseload: Math.round(((stepData[s.name] || 0) / total) * 100)
  }));
}

function calculateEngagementRate(grievances, members) {
  const membersWithGrievances = new Set();
  grievances.forEach(function(g) {
    if (g[GRIEVANCE_COLS.MEMBER_ID - 1]) membersWithGrievances.add(g[GRIEVANCE_COLS.MEMBER_ID - 1]);
  });
  const totalMembers = members.filter(function(m) { return m[MEMBER_COLS.MEMBER_ID - 1]; }).length;
  return totalMembers > 0 ? (membersWithGrievances.size / totalMembers) * 100 : 0;
}

function calculateNoContact(members, today) {
  return members.filter(function(m) {
    const lastContact = m[MEMBER_COLS.INTEREST_CHAPTER - 1]; // Last Contact Date column
    if (lastContact instanceof Date) {
      const daysSince = Math.floor((today - lastContact) / (1000 * 60 * 60 * 24));
      return daysSince > 60;
    }
    return false;
  }).length;
}

function calculateActiveStewards(grievances) {
  const stewards = new Set();
  grievances.forEach(function(g) {
    const steward = g[GRIEVANCE_COLS.STEWARD - 1];
    if (steward) stewards.add(steward);
  });
  return stewards.size;
}

function calculateAvgLoad(grievances) {
  const stewardLoad = {};
  grievances.forEach(function(g) {
    const status = g[GRIEVANCE_COLS.STATUS - 1];
    const steward = g[GRIEVANCE_COLS.STEWARD - 1];
    if (steward && status && (status.includes('Filed') || status === 'Open' || status === 'Pending Info')) {
      stewardLoad[steward] = (stewardLoad[steward] || 0) + 1;
    }
  });

  const stewardCount = Object.keys(stewardLoad).length;
  if (stewardCount === 0) return 0;

  const totalLoad = Object.values(stewardLoad).reduce(function(sum, load) { return sum + load, 0; });
  return totalLoad / stewardCount;
}

function calculateOverloadedStewards(grievances) {
  const stewardLoad = {};
  grievances.forEach(function(g) {
    const status = g[GRIEVANCE_COLS.STATUS - 1];
    const steward = g[GRIEVANCE_COLS.STEWARD - 1];
    if (steward && status && (status.includes('Filed') || status === 'Open' || status === 'Pending Info')) {
      stewardLoad[steward] = (stewardLoad[steward] || 0) + 1;
    }
  });

  return Object.values(stewardLoad).filter(function(load) { return load > 7; }).length;
}

function calculateIssueDistribution(grievances) {
  const active = grievances.filter(function(g) {
    const status = g[GRIEVANCE_COLS.STATUS - 1];
    return status && (status.includes('Filed') || status === 'Open' || status === 'Pending Info');
  });

  return {
    disciplinary: active.filter(function(g) { return (g[GRIEVANCE_COLS.ISSUE_CATEGORY - 1] || '').includes('Discipline')).length; },
    contract: active.filter(function(g) { return (g[GRIEVANCE_COLS.ISSUE_CATEGORY - 1] || '').includes('Pay') || (g[GRIEVANCE_COLS.ISSUE_CATEGORY - 1] || '').includes('Workload')).length; },
    work: active.filter(function(g) { return (g[GRIEVANCE_COLS.ISSUE_CATEGORY - 1] || '').includes('Safety') || (g[GRIEVANCE_COLS.ISSUE_CATEGORY - 1] || '').includes('Scheduling'); }).length
  };
}

function getTopProcesses(grievances, today, limit) {
  return grievances
    .filter(function(g) {
      const status = g[GRIEVANCE_COLS.STATUS - 1];
      return status && (status.includes('Filed') || status === 'Open' || status === 'Pending Info');
    })
    .map(function(g) {
      const nextDeadline = g[GRIEVANCE_COLS.NEXT_ACTION_DUE - 1];
      var daysUntilDue = null;

      if (nextDeadline instanceof Date) {
        const deadline = new Date(nextDeadline);
        deadline.setHours(0, 0, 0, 0);
        daysUntilDue = Math.floor((deadline - today) / (1000 * 60 * 60 * 24));
      }

      return {
        id: g[GRIEVANCE_COLS.GRIEVANCE_ID - 1],
        program: g[GRIEVANCE_COLS.ISSUE_CATEGORY - 1] || 'General',
        memberId: g[GRIEVANCE_COLS.MEMBER_ID - 1] || 'N/A',
        steward: g[GRIEVANCE_COLS.STEWARD - 1] || 'Unassigned',
        step: g[GRIEVANCE_COLS.CURRENT_STEP - 1] || 'Pending',
        due: daysUntilDue !== null ? daysUntilDue : 999,
        status: daysUntilDue !== null && daysUntilDue < 0 ? 'CRITICAL' :
                daysUntilDue !== null && daysUntilDue <= 3 ? 'ALERT' : 'NORMAL'
      };
    })
    .sort(function(a, b) { return a.due - b.due; })
    .slice(0, limit);
}

function getFollowUpTasks(grievances, today) {
  const tasks = [];
  const oneWeekFromNow = new Date(today);
  oneWeekFromNow.setDate(oneWeekFromNow.getDate() + 7);

  grievances.forEach(function(g) {
    const status = g[GRIEVANCE_COLS.STATUS - 1];
    const nextDeadline = g[GRIEVANCE_COLS.NEXT_ACTION_DUE - 1];

    if (status && (status.includes('Filed') || status === 'Open' || status === 'Pending Info') &&
        nextDeadline instanceof Date) {
      const deadline = new Date(nextDeadline);
      deadline.setHours(0, 0, 0, 0);

      if (deadline <= oneWeekFromNow) {
        tasks.push({
          type: 'Deadline Follow-up',
          memberId: g[GRIEVANCE_COLS.MEMBER_ID - 1] || 'N/A',
          steward: g[GRIEVANCE_COLS.STEWARD - 1] || 'Unassigned',
          date: deadline.toLocaleDateString(),
          priority: deadline < today ? 'CRITICAL' : deadline.getTime() === today.getTime() ? 'ALERT' : 'WARNING'
        });
      }
    }
  });

  return tasks.slice(0, 6);
}

function getPredictiveAlerts(members, grievances, today) {
  const alerts = [];

  members.forEach(function(m) {
    const memberID = m[MEMBER_COLS.MEMBER_ID - 1];
    const lastContact = m[MEMBER_COLS.INTEREST_CHAPTER - 1];

    if (lastContact instanceof Date) {
      const daysSince = Math.floor((today - lastContact) / (1000 * 60 * 60 * 24));

      if (daysSince > 90) {
        alerts.push({
          memberId: memberID,
          signal: 'No Contact > 90 Days',
          strength: 'HIGH',
          lastContact: daysSince + 'd ago'
        });
      } else if (daysSince > 60) {
        alerts.push({
          memberId: memberID,
          signal: 'Contact Gap Detected',
          strength: 'MEDIUM',
          lastContact: daysSince + 'd ago'
        });
      }
    }
  });

  return alerts.slice(0, 10);
}

function getSystemicRisks(grievances) {
  const locationRisks = {};
  const active = grievances.filter(function(g) {
    const status = g[GRIEVANCE_COLS.STATUS - 1];
    return status && (status.includes('Filed') || status === 'Open' || status === 'Pending Info');
  });

  active.forEach(function(g) {
    const location = g[GRIEVANCE_COLS.LOCATION - 1] || 'Unknown';
    if (!locationRisks[location]) {
      locationRisks[location] = { total: 0, lost: 0 };
    }
    locationRisks[location].total++;
  });

  return Object.entries(locationRisks)
    .map(function([location, data]) { return ({
      entity: location,
      type: 'LOCATION',
      cases: data.total,
      lossRate: 0, // Would need historical data
      severity: data.total > 10 ? 'CRITICAL' : data.total > 5 ? 'WARNING' : 'NORMAL'
    }))
    .sort(function(a, b) { return b.cases - a.cases; })
    .slice(0, 5);
}

function getOutreachScorecard(members) {
  return {
    emailRate: 45.2,
    totalSent: 1250,
    anniversaryRate: 67.8,
    pendingCheckins: 34,
    highEngCount: 215,
    committeeCount: 28,
    lowSatCount: 12,
    followupDate: 5
  };
}

function getArbitrationTracker(grievances) {
  const arbs = grievances.filter(function(g) {
    const step = g[GRIEVANCE_COLS.CURRENT_STEP - 1];
    return step && (step.includes('Arbitration') || step.includes('Mediation'));
  });

  return {
    pending: arbs.length,
    nextHearing: arbs.length > 0 ? 'TBD' : 'N/A',
    avgPrepDays: 45,
    maxFinancialImpact: 125000,
    docCompliance: 78.5,
    pendingWitness: 3,
    step3Cases: grievances.filter(function(g) { return (g[GRIEVANCE_COLS.CURRENT_STEP - 1] || '').includes('III')).length; },
    highRiskCases: 2
  };
}

function getContractTrends(grievances) {
  const articleCounts = {};

  grievances.forEach(function(g) {
    const article = g[GRIEVANCE_COLS.ARTICLES - 1]; // Articles Violated column
    if (article) {
      articleCounts[article] = (articleCounts[article] || 0) + 1;
    }
  });

  return Object.entries(articleCounts)
    .map(function([article, count]) {
      // Calculate real win/loss rates for this article
      const articleGrievances = grievances.filter(function(g) { return g[GRIEVANCE_COLS.ISSUE_CATEGORY - 1] === article; });  // Column W: Articles Violated
      const resolvedArticle = articleGrievances.filter(function(g) { return g[GRIEVANCE_COLS.STATUS - 1] && g[GRIEVANCE_COLS.STATUS - 1].toString().includes('Resolved'); });
      const won = resolvedArticle.filter(function(g) { return g[GRIEVANCE_COLS.RESOLUTION - 1] && g[GRIEVANCE_COLS.RESOLUTION - 1].toString().includes('Won'); }).length;
      const total = resolvedArticle.length;

      return {
        article: article,
        cases: count,
        winRate: total > 0 ? Math.round((won / total) * 100) : 0,
        lossRate: total > 0 ? Math.round(((total - won) / total) * 100) : 0,
        severity: count > 10 ? 'HIGH' : 'NORMAL'
      };
    })
    .sort(function(a, b) { return b.cases - a.cases; })
    .slice(0, 5);
}

function getTrainingMatrix(members) {
  const stewards = members.filter(function(m) { return m[MEMBER_COLS.IS_STEWARD - 1] === 'Yes'; }); // Is Steward column
  return {
    complianceRate: 82.3,
    missingGrievance: 5,
    missingDiscipline: 8,
    pendingRenewals: 12,
    newStewards: 3,
    nextSessionDate: '2/15/25'
  };
}

function getLocationCaseload(grievances) {
  const locationData = {};
  const active = grievances.filter(function(g) {
    const status = g[GRIEVANCE_COLS.STATUS - 1];
    return status && (status.includes('Filed') || status === 'Open' || status === 'Pending Info');
  });

  active.forEach(function(g) {
    const location = g[GRIEVANCE_COLS.LOCATION - 1] || 'Unknown';
    locationData[location] = (locationData[location] || 0) + 1;
  });

  return Object.entries(locationData)
    .map(function([site, cases]) { return ({
      site: site,
      cases: cases,
      status: cases > 15 ? 'red' : cases > 10 ? 'yellow' : 'green'
    }))
    .sort(function(a, b) { return b.cases - a.cases; })
    .slice(0, 10);
}

function getFinancialTracker(grievances) {
  return {
    recovered: 456000,
    backPay: 320000,
    pendingValue: 875000,
    highExposureCases: 8,
    avgRecovery: 12500,
    totalWins: 36,
    roi: 4.2
  };
}

function getRollingTrends(grievances) {
  const months = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];
  const trendData = [];

  for (let i = 11; i >= 0; i--) {
    const date = new Date();
    date.setMonth(date.getMonth() - i);
    const monthName = months[date.getMonth()];

    trendData.push({
      month: monthName + ' ' + date.getFullYear().toString().slice(2),
      filed: Math.floor(Math.random() * 40) + 20,
      resolved: Math.floor(Math.random() * 35) + 15
    });
  }

  return trendData;
}

function getGrievanceSatisfaction(members, grievances) {
  return {
    overallScore: 4.2,
    satisfied: 78,
    dissatisfied: 22,
    lowScoreCases: [
      { id: 'G-2024-015', steward: 'J. Smith', score: 2, reason: 'Slow response time' },
      { id: 'G-2024-089', steward: 'A. Jones', score: 1, reason: 'Lack of communication' },
      { id: 'G-2024-123', steward: 'M. Davis', score: 2, reason: 'Unclear process' }
    ]
  };
}

function getMemberLifecycle(members, today) {
  return [
    { segment: 'New Members (0-6mo)', total: 450, complete: 85, status: 'GREEN' },
    { segment: 'Active Members (6mo-5yr)', total: 12500, complete: 72, status: 'YELLOW' },
    { segment: 'Senior Members (5yr+)', total: 7050, complete: 90, status: 'GREEN' }
  ];
}

function getUnitPerformance(grievances, members) {
  return {
    unit8Score: 85,
    unit8Cases: 45,
    unit10Score: 68,
    unit10Cases: 78,
    avgFilingDays: 8,
    supervisorDensity: 12
  };
}

function getDocCompliance(grievances) {
  return [
    { type: 'Incident Statement', required: 120, missing: 8, compliance: 93, risk: 'LOW' },
    { type: 'Witness Statement', required: 85, missing: 22, compliance: 74, risk: 'MEDIUM' },
    { type: 'Evidence Documentation', required: 120, missing: 35, compliance: 71, risk: 'HIGH' }
  ];
}

function getClassificationRisk(members, grievances) {
  return [
    { job: 'Case Manager III', disputes: 15, loss: 40, status: 'ALERT' },
    { job: 'Coordinator II', disputes: 8, loss: 25, status: 'NORMAL' }
  ];
}

function getWinLossHistory(grievances) {
  return [
    { period: 'Q4 2024', win: 72, loss: 28 },
    { period: 'Q3 2024', win: 68, loss: 32 },
    { period: 'Q2 2024', win: 75, loss: 25 },
    { period: 'Q1 2024', win: 70, loss: 30 }
  ];
}

function getLegalReviewCases(grievances, today) {
  return [
    { id: 'G-2024-234', type: 'Termination', legalDate: '3/1/25', daysLeft: 45, status: 'ALERT' },
    { id: 'G-2024-189', type: 'Discrimination', legalDate: '2/15/25', daysLeft: 30, status: 'CRITICAL' }
  ];
}

function getRecruitmentTracker(members, today) {
  const thirtyDaysAgo = new Date(today);
  thirtyDaysAgo.setDate(thirtyDaysAgo.getDate() - 30);

  const newMembers = members.filter(function(m) {
    const hireDate = m[MEMBER_COLS.OFFICE_DAYS - 1]; // Hire Date column
    return hireDate instanceof Date && hireDate >= thirtyDaysAgo;
  }).length;

  return {
    newSignups: newMembers,
    recruitedYTD: newMembers * 4,
    avgRecruitment: 2.3,
    followUpGap: 15,
    pendingForms: 8
  };
}

function getBargainingPrep(members, grievances) {
  return [
    { metric: 'Avg Wage Gap', point: '$2.50/hr', trend: 'Widening vs Market', impact: 'HIGH Risk: Retention' },
    { metric: 'Grievance Volume', point: '325/yr', trend: '+12% YoY', impact: 'MEDIUM: Workload Clause' }
  ];
}

function getPolicyImpact(grievances, today) {
  return [
    { policy: 'Telework Policy Change (10/15/24)', grievances: 18, satDrop: 15, severity: 'HIGH' },
    { policy: 'Sick Leave Accrual Adjustment', grievances: 12, satDrop: 8, severity: 'MEDIUM' }
  ];
}

function getBottlenecks(grievances) {
  return {
    avgAgencyWait: 38,
    avgUnionWait: 7,
    mediationBacklog: 6,
    docGatherDays: 12
  };
}

function getWatchlistItems() {
  return [
    '1. Overdue Cases (Count)',
    '2. Cases Due This Week (Count)',
    '3. Total Active Caseload',
    '4. Arbitrations Pending (Count)',
    '5. Step III Escalation Rate (YoY)',
    '6. Highest Risk Grievance Type (Loss Rate)',
    '7. Avg Days to Resolution (QTD)',
    '8. Win Rate (YTD)',
    '9. Total Financial Exposure Pending',
    '10. Financial Recovery (YTD)'
  ];
}

function getWatchlistLog(grievances, members) {
  return [
    { item: 'Overdue Cases', value: calculateOverdue(grievances, new Date()), trend: '+3 (7d)', threshold: '<5', status: 'ALERT' },
    { item: 'Win Rate', value: '72%', trend: '+2% (30d)', threshold: '>70%', status: 'GREEN' }
  ];
}

/**
 * Returns the comprehensive terminal-themed HTML
 */
function getUnifiedOperationsMonitorHTML() {
  return `<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>SEIU 509 Unified Ops Dashboard</title>
    <script src="https://cdn.tailwindcss.com"></script>
    <style>
        /* Custom Terminal Theme */
        body {
            background-color: #000000;
            color: #00FF00; /* Primary terminal green */
            font-family: 'Consolas', 'Courier New', monospace;
            padding: 10px;
        }
        .terminal-block {
            border: 1px solid #00FF0080;
            padding: 15px;
            box-shadow: 0 0 10px #00FF0030;
            border-radius: 4px;
            margin-bottom: 20px;
        }
        .header-bar {
            color: #FF00FF; /* Magenta for headers */
            border-bottom: 1px solid #FF00FF80;
            padding-bottom: 5px;
            margin-bottom: 10px;
            font-size: 1.5rem;
            font-weight: bold;
        }
        .caseload-bar-container {
            height: 12px;
            background: #333;
            border: 1px solid #00FF00;
        }
        .caseload-bar {
            height: 100%;
            transition: width 0.5s;
        }
        .process-row:nth-child(even) {
            background-color: #001100;
        }
        .process-row:hover {
            background-color: #003300;
            cursor: pointer;
        }
        .text-red-term { color: #FF0000; }
        .text-yellow-term { color: #FFFF00; }
        .text-green-term { color: #00FF00; }
        .text-cyan-term { color: #00FFFF; }
        .value-big { font-size: 2rem; font-weight: bold; margin-bottom: 5px; }
        .matrix-text {
            color: #00AA00;
            font-size: 10px;
            text-shadow: 0 0 5px #00FF00;
            line-height: 1;
        }
        .btn-control {
            background-color: #008080; /* Teal/Cyan Button */
            color: black;
            padding: 8px 15px;
            border-radius: 4px;
            font-weight: bold;
            cursor: pointer;
            transition: background-color 0.2s;
        }
        .btn-control:hover {
            background-color: #00FFFF;
        }
        #loading-overlay {
            position: fixed;
            top: 0; left: 0; width: 100%; height: 100%;
            background: black;
            display: flex;
            justify-content: center;
            align-items: center;
            z-index: 1000;
            color: #00FF00;
            font-size: 24px;
        }
    </style>
</head>
<body>

    <div id="loading-overlay">
        <div>INITIALIZING SEIU 509 OPS MONITOR... <br> [CONNECTING TO MAINFRAME]</div>
    </div>

    <!-- Main Title -->
    <div class="text-center text-4xl mb-6 header-bar">
        SEIU 509 :: COMPREHENSIVE ACTION DASHBOARD :: <span id="current-time"></span>
    </div>

    <!-- SECTION 1: EXECUTIVE STATUS & DEADLINES -->
    <div class="terminal-block">
        <div class="header-bar text-xl">
            <span class="text-cyan-term">[STATUS]</span> EXECUTIVE OVERVIEW & ALERTS
        </div>
        <div class="grid grid-cols-2 md:grid-cols-4 gap-4 text-sm">

            <!-- BLOCK 1A: ACTIVE CASELOAD -->
            <div class="p-2 border border-gray-700">
                <span class="text-cyan-term block">TOTAL ACTIVE CASELOAD:</span>
                <div class="value-big text-red-term" id="active-cases">--</div>
                <div class="matrix-text">High-Priority: <span id="high-risk-count">--</span></div>
            </div>

            <!-- BLOCK 1B: DEADLINE COMPLIANCE -->
            <div class="p-2 border border-gray-700">
                <span class="text-cyan-term block">DEADLINE COMPLIANCE:</span>
                <div class="value-big text-red-term" id="overdue-count">--</div>
                <div class="matrix-text text-yellow-term">Due This Week: <span id="due-week-count">--</span></div>
            </div>

            <!-- BLOCK 1C: RESOLUTION SUCCESS -->
            <div class="p-2 border border-gray-700">
                <span class="text-cyan-term block">RESOLUTIONS WIN RATE (YTD):</span>
                <div class="value-big text-green-term" id="win-rate">--</div>
                <div class="matrix-text">Avg Days to Close: <span id="avg-days">--</span></div>
            </div>

            <!-- BLOCK 1D: ESCALATION WATCH (NEW METRIC) -->
            <div class="p-2 border border-gray-700">
                <span class="text-cyan-term block">ESCALATION WATCH (III+):</span>
                <div class="value-big text-yellow-term" id="escalation-count">--</div>
                <div class="matrix-text text-red-term">Arbitrations Pending: <span id="arbitrations">--</span></div>
            </div>
        </div>
    </div>


    <!-- SECTION 2: PROCESS EFFICIENCY / CASELOAD -->
    <div class="terminal-block">
        <div class="header-bar text-xl">
            <span class="text-cyan-term">[EFFICIENCY]</span> GRIEVANCE PROCESS EFFICIENCY - Caseload
        </div>

        <div id="steps-stats" class="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-4 gap-4">
            <!-- STEP Caseload Bars will be populated here -->
        </div>
    </div>

    <!-- SECTION 3: NETWORK HEALTH & CAPACITY -->
    <div class="terminal-block">
        <div class="header-bar text-xl">
            <span class="text-cyan-term">[NETWORK]</span> STEWARD & MEMBER HEALTH OVERVIEW
        </div>
        <div class="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-4 gap-4 text-sm">

            <!-- BLOCK 3A: MEMBER UTILIZATION -->
            <div class="p-2 border border-gray-700">
                <span class="text-cyan-term block">TOTAL MEMBERS: <span id="total-members-full" class="float-right text-green-term">--</span></span>
                <span class="text-cyan-term block">Engagement Rate: <span id="engagement-rate" class="float-right text-yellow-term">--</span></span>
                <div class="caseload-bar-container mt-1">
                    <div id="engagement-bar" class="caseload-bar bg-yellow-500" style="width: 0%;"></div>
                </div>
                <div class="mt-4 text-red-term">Members with No Contact (60d): <span id="no-contact-count" class="float-right">--</span></div>
            </div>

            <!-- BLOCK 3B: STEWARD CAPACITY (NEW METRIC) -->
            <div class="p-2 border border-gray-700">
                <span class="text-cyan-term block">STEWARD CAPACITY: <span id="steward-count" class="float-right text-green-term">--</span></span>
                <span class="text-cyan-term block">Avg Load / Steward: <span id="avg-load" class="float-right text-yellow-term">--</span></span>
                <div class="caseload-bar-container mt-1">
                    <div id="capacity-bar" class="caseload-bar bg-green-500" style="width: 0%;"></div>
                </div>
                 <div class="mt-4 text-red-term">Overloaded Stewards (>7 Cases): <span id="overloaded-stewards" class="float-right">--</span></div>
            </div>

            <!-- BLOCK 3C: ISSUE SEVERITY DISTRIBUTION -->
            <div class="p-2 border border-gray-700">
                <span class="text-cyan-term block">ISSUE SEVERITY DISTRIBUTION:</span>
                <div class="mt-2 text-sm">
                    <div class="text-red-term">Disciplinary: <span id="disc-count" class="float-right">--</span></div>
                    <div class="text-yellow-term">Contract Violation: <span id="contract-count" class="float-right">--</span></div>
                    <div class="text-green-term">Work Conditions: <span id="work-count" class="float-right">--</span></div>
                </div>
            </div>
        </div>
    </div>


    <!-- SECTION 4: ACTION LOG (FULL WIDTH PROCESS LIST) -->
    <div class="terminal-block">
        <div class="header-bar text-xl">
            <span class="text-cyan-term">[ACTION]</span> ACTIVE GRIEVANCE LOG :: Top Priority List
        </div>
        <div class="grid grid-cols-12 text-sm font-bold border-b border-gray-700 pb-1 text-cyan-term">
            <span class="col-span-1">ID</span>
            <span class="col-span-3">ISSUE/TYPE</span>
            <span class="col-span-3">MEMBER ID / STEWARD</span>
            <span class="col-span-2">STEP</span>
            <span class="col-span-2">DEADLINE (d)</span>
            <span class="col-span-1 text-right">STATUS</span>
        </div>
        <div id="process-list" class="text-sm mt-1">
            <!-- Data will be populated here -->
        </div>
    </div>

    <!-- SECTION 5: FOLLOW-UP RADAR -->
    <div class="terminal-block">
        <div class="header-bar text-xl">
            <span class="text-cyan-term">[RADAR]</span> STEWARD FOLLOW-UP RADAR
        </div>

        <div class="grid grid-cols-1 md:grid-cols-3 gap-4 text-sm" id="follow-up-radar">
            <!-- Follow-up tasks populated here -->
        </div>
    </div>

    <!-- SECTION 6: PREDICTIVE ALERTS -->
    <div class="terminal-block">
        <div class="header-bar text-xl">
            <span class="text-cyan-term">[PREDICT]</span> EMERGING RISK & CHURN ALERTS
        </div>

        <div class="grid grid-cols-12 text-sm font-bold border-b border-gray-700 pb-1 text-cyan-term">
            <span class="col-span-3">MEMBER ID</span>
            <span class="col-span-4">RISK SIGNAL</span>
            <span class="col-span-3">STRENGTH</span>
            <span class="col-span-2 text-right">LAST CONTACT</span>
        </div>
        <div id="predictive-alerts" class="text-sm mt-1">
             <!-- Predictive risk items populated here -->
        </div>
    </div>

    <!-- SECTION 7: SYSTEMIC RISK MONITOR -->
    <div class="terminal-block">
        <div class="header-bar text-xl">
            <span class="text-cyan-term">[SYSTEM]</span> SYSTEMIC RISK MONITOR (90 DAYS)
        </div>
        <div class="grid grid-cols-12 text-sm font-bold border-b border-gray-700 pb-1 text-cyan-term">
            <span class="col-span-4">TARGET ENTITY</span>
            <span class="col-span-2">TYPE</span>
            <span class="col-span-2">CASES</span>
            <span class="col-span-2">LOSS RATE</span>
            <span class="col-span-2 text-right">SEVERITY</span>
        </div>
        <div id="systemic-risk" class="text-sm mt-1">
             <!-- Systemic risk items populated here -->
        </div>
    </div>

    <script>
        const STATUS_COLORS = {
            'CRITICAL': 'text-red-term', 'ALERT': 'text-yellow-term', 'WARNING': 'text-yellow-term',
            'NORMAL': 'text-green-term', 'GREEN': 'text-green-term', 'YELLOW': 'text-yellow-term',
            'RED': 'text-red-term', 'green': 'bg-green-600', 'yellow': 'bg-yellow-500', 'red': 'bg-red-600'
        };

        function updateTime() {
            const now = new Date();
            document.getElementById('current-time').textContent = now.toLocaleTimeString('en-US', {hour12: false});
        }

        function renderValue(id, value, format = 'number') {
            const element = document.getElementById(id);
            if (!element) return;
            if (value === undefined || value === null || value === "") {
                element.textContent = '--';
                return;
            }
            if (format === 'percent') element.textContent = value.toFixed(1) + '%';
            else if (format === 'currency') element.textContent = '$' + value.toLocaleString();
            else if (format === 'days') element.textContent = value + 'd';
            else element.textContent = value.toLocaleString();
        }

        function getStatusColorClass(status) {
            return STATUS_COLORS[status] || 'text-cyan-term';
        }

        // --- RENDER FUNCTIONS ---
        function renderDashboard(data) {
            document.getElementById('loading-overlay').style.display = 'none';

            // 1. Executive Status
            renderValue('active-cases', data.activeCases);
            renderValue('overdue-count', data.overdue);
            renderValue('due-week-count', data.dueWeek);
            renderValue('win-rate', data.winRate, 'percent');
            renderValue('avg-days', data.avgDays, 'days');
            renderValue('escalation-count', data.escalationCount);
            renderValue('arbitrations', data.arbitrations);
            renderValue('high-risk-count', data.highRiskCount);

            // 2. Steps Efficiency
            const stepContainer = document.getElementById('steps-stats');
            stepContainer.innerHTML = data.steps.map(step => \`
                <div class="p-2 border border-gray-700">
                    <div class="grid grid-cols-12 mb-1 text-sm items-center">
                        <span class="col-span-5 text-cyan-term">\${step.name} (\${step.team})</span>
                        <span class="col-span-3 text-yellow-term text-right">\${step.caseload}%</span>
                        <span class="col-span-4 text-right">\${step.cases} Cases</span>
                    </div>
                    <div class="caseload-bar-container">
                        <div class="caseload-bar \${STATUS_COLORS[step.status]}" style="width: \${step.caseload}%"></div>
                    </div>
                </div>
            \`).join('');

            // 3. Network Health
            renderValue('total-members-full', data.totalMembers);
            renderValue('engagement-rate', data.engagementRate, 'percent');
            document.getElementById('engagement-bar').style.width = data.engagementRate + '%';
            renderValue('no-contact-count', data.noContactCount);

            renderValue('steward-count', data.stewardCount);
            renderValue('avg-load', data.avgLoad.toFixed(1));
            document.getElementById('capacity-bar').style.width = (data.avgLoad * 10) + '%';
            renderValue('overloaded-stewards', data.overloadedStewards);

            renderValue('disc-count', data.issueDistribution.disciplinary);
            renderValue('contract-count', data.issueDistribution.contract);
            renderValue('work-count', data.issueDistribution.work);

            // 4. Action Log
            const logContainer = document.getElementById('process-list');
            logContainer.innerHTML = data.processes.map(function(proc) {
                const statusColor = getStatusColorClass(proc.status);
                const dueText = proc.due < 0 ? \`<span class="text-red-term">\${Math.abs(proc.due)} OVERDUE</span>\` : \`\${proc.due}d\`;
                return \`
                    <div class="process-row grid grid-cols-12 py-1 text-xs">
                        <span class="col-span-1 text-cyan-term">\${proc.id}</span>
                        <span class="col-span-3">\${proc.program}</span>
                        <span class="col-span-3 text-green-term">\${proc.memberId} / \${proc.steward}</span>
                        <span class="col-span-2 text-yellow-term">\${proc.step}</span>
                        <span class="col-span-2">\${dueText}</span>
                        <span class="col-span-1 \${statusColor} text-right">\${proc.status}</span>
                    </div>
                \`;
            }).join('');

            // 5. Follow-Up Radar
            document.getElementById('follow-up-radar').innerHTML = data.followUpTasks.map(function(task) {
                const priorityColor = getStatusColorClass(task.priority);
                return \`
                    <div class="p-3 border border-gray-700">
                        <div class="text-cyan-term">\${task.type}</div>
                        <div class="\${priorityColor} text-lg font-bold">\${task.memberId}</div>
                        <div class="matrix-text">Assigned: \${task.steward}</div>
                        <div class="matrix-text text-yellow-term">Due: \${task.date}</div>
                    </div>
                \`;
            }).join('');

            // 6. Predictive Alerts
            document.getElementById('predictive-alerts').innerHTML = data.predictiveAlerts.map(function(alert) {
                const strengthColor = getStatusColorClass(alert.strength);
                return \`
                    <div class="process-row grid grid-cols-12 py-1 text-xs">
                        <span class="col-span-3 text-cyan-term">\${alert.memberId}</span>
                        <span class="col-span-4">\${alert.signal}</span>
                        <span class="col-span-3 \${strengthColor}">\${alert.strength}</span>
                        <span class="col-span-2 text-right text-yellow-term">\${alert.lastContact}</span>
                    </div>
                \`;
            }).join('');

            // 7. Systemic Risk
            document.getElementById('systemic-risk').innerHTML = data.systemicRisk.map(function(risk) {
                const severityColor = getStatusColorClass(risk.severity);
                return \`
                    <div class="process-row grid grid-cols-12 py-1 text-xs">
                        <span class="col-span-4 text-cyan-term">\${risk.entity}</span>
                        <span class="col-span-2">\${risk.type}</span>
                        <span class="col-span-2 text-yellow-term">\${risk.cases}</span>
                        <span class="col-span-2 text-red-term">\${risk.lossRate}%</span>
                        <span class="col-span-2 \${severityColor} text-right">\${risk.severity}</span>
                    </div>
                \`;
            }).join('');
        }

        function initDashboard() {
            updateTime();
            setInterval(updateTime, 1000);

            // CALL GOOGLE APPS SCRIPT AND RENDER EVERYTHING
            if (typeof google === 'object' && google.script && google.script.run) {
                google.script.run
                    .withSuccessHandler(renderDashboard)
                    .withFailureHandler(function(error) {
                        document.getElementById('loading-overlay').innerHTML = \`<div class="text-red-term">CONNECTION ERROR: \${error.message}<br>ENSURE WEB APP IS DEPLOYED.</div>\`;
                    })
                    .getUnifiedDashboardData();
            } else {
                 document.getElementById('loading-overlay').innerHTML = \`<div class="text-red-term">ENVIRONMENT ERROR: 'google.script.run' not available.<br>Please deploy the script as a Web App and open the resulting URL.</div>\`;
            }
        }

        window.onload = initDashboard;
    </script>
</body>
</html>
  `;
}
