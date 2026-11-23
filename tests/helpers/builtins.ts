/* Test fixture: deterministic, domain-themed problem drafts (the app itself always generates problems with Gemini). */
import type { ColumnType, ConceptId, Domain, ProblemDraft, ProblemLevel } from "../../src/mentor/types"
import { rng } from "../../src/mentor/random"

interface Theme {
  title: string
  role: string
  entity: string
  entityPlural: string
  nameCol: string
  names: string[]
  periods: string[]
  measure: string
  type: ColumnType
  range: [number, number]
  step: number
  category: string
  categories: string[]
  targetCol: string
  manager: string
  lookupCol: string
  currency?: string
}

const PEOPLE = ['Rahul Sharma', 'Priya Nair', 'Arjun Mehta', 'Sneha Iyer', 'Vikram Rao', 'Ananya Gupta', 'Karan Singh', 'Meera Pillai', 'Rohan Das', 'Fatima Khan', 'David Wilson', 'Aisha Patel', 'Sanjay Kumar', 'Neha Joshi', 'Arun Menon', 'Divya Reddy', 'Imran Shaikh', 'Lakshmi Varma', 'Tom Baker', 'Kavya Bhat']

const THEMES: Record<Domain, Theme> = {
  Sales: { title: 'Sales Performance Review', role: 'Sales Analyst', entity: 'sales representative', entityPlural: 'sales representatives', nameCol: 'Sales Rep', names: PEOPLE, periods: ['Jan', 'Feb', 'Mar', 'Apr'], measure: 'sales', type: 'currency', range: [18000, 62000], step: 500, category: 'Region', categories: ['North', 'South', 'East', 'West'], targetCol: 'Target', manager: 'sales manager', lookupCol: 'Commission Rate', currency: '$' },
  HR: { title: 'Employee Working Hours Review', role: 'HR Executive', entity: 'employee', entityPlural: 'employees', nameCol: 'Employee', names: PEOPLE, periods: ['Week 1', 'Week 2', 'Week 3', 'Week 4'], measure: 'hours worked', type: 'integer', range: [32, 48], step: 1, category: 'Department', categories: ['Finance', 'HR', 'Sales', 'IT'], targetCol: 'Required Hours', manager: 'HR manager', lookupCol: 'Overtime Rate' },
  Finance: { title: 'Department Expense Tracking', role: 'Finance Analyst', entity: 'cost centre', entityPlural: 'cost centres', nameCol: 'Cost Centre', names: ['Admin', 'Facilities', 'IT Support', 'Marketing Ops', 'Payroll', 'Legal', 'Procurement', 'Training', 'Travel Desk', 'Security', 'Research', 'Customer Care', 'Logistics', 'Quality', 'Compliance', 'Events', 'Analytics', 'Internal Audit'], periods: ['Q1', 'Q2', 'Q3', 'Q4'], measure: 'expenses', type: 'currency', range: [12000, 48000], step: 250, category: 'Division', categories: ['Corporate', 'Operations', 'Support'], targetCol: 'Budget', manager: 'finance controller', lookupCol: 'Overhead Rate', currency: '$' },
  Marketing: { title: 'Campaign Lead Analysis', role: 'Marketing Analyst', entity: 'campaign', entityPlural: 'campaigns', nameCol: 'Campaign', names: ['Spring Launch', 'Summer Deals', 'Webinar Series', 'Referral Push', 'Festive Offer', 'Product Demo', 'Newsletter', 'Retargeting', 'Brand Video', 'Student Promo', 'Partner Week', 'App Install', 'Loyalty Drive', 'Flash Sale', 'Trade Show', 'Podcast Ads', 'Influencer Mix', 'Local Search'], periods: ['Week 1', 'Week 2', 'Week 3', 'Week 4'], measure: 'leads', type: 'integer', range: [40, 260], step: 1, category: 'Channel', categories: ['Email', 'Social', 'Search', 'Events'], targetCol: 'Lead Goal', manager: 'marketing head', lookupCol: 'Cost per Lead' },
  Operations: { title: 'Warehouse Dispatch Report', role: 'Operations Analyst', entity: 'product', entityPlural: 'products', nameCol: 'Product', names: ['Steel Bolts', 'Copper Wire', 'LED Panels', 'Paint Cans', 'Safety Gloves', 'PVC Pipes', 'Cable Ties', 'Drill Bits', 'Wall Plugs', 'Glue Sticks', 'Hinges', 'Door Locks', 'Switch Boards', 'Ceiling Fans', 'Water Pumps', 'Tool Kits', 'Ladders', 'Tape Rolls'], periods: ['Jan', 'Feb', 'Mar', 'Apr'], measure: 'units dispatched', type: 'integer', range: [120, 980], step: 5, category: 'Category', categories: ['Hardware', 'Electrical', 'Plumbing'], targetCol: 'Dispatch Target', manager: 'operations manager', lookupCol: 'Handling Rate' },
  Healthcare: { title: 'Clinic Appointment Analysis', role: 'Healthcare Analyst', entity: 'doctor', entityPlural: 'doctors', nameCol: 'Doctor', names: PEOPLE.map((n) => `Dr. ${n}`), periods: ['Week 1', 'Week 2', 'Week 3', 'Week 4'], measure: 'appointments', type: 'integer', range: [25, 90], step: 1, category: 'Department', categories: ['Cardiology', 'Pediatrics', 'Orthopedics', 'General'], targetCol: 'Capacity', manager: 'clinic administrator', lookupCol: 'Fee Share' },
  Education: { title: 'Student Marks Review', role: 'Academic Coordinator', entity: 'student', entityPlural: 'students', nameCol: 'Student', names: PEOPLE, periods: ['Test 1', 'Test 2', 'Test 3', 'Test 4'], measure: 'marks', type: 'integer', range: [35, 98], step: 1, category: 'Section', categories: ['A', 'B', 'C'], targetCol: 'Pass Mark', manager: 'head of department', lookupCol: 'Scholarship %' },
  'Customer Support': { title: 'Support Ticket Resolution Report', role: 'Support Team Lead', entity: 'agent', entityPlural: 'agents', nameCol: 'Agent', names: PEOPLE, periods: ['Week 1', 'Week 2', 'Week 3', 'Week 4'], measure: 'tickets resolved', type: 'integer', range: [30, 120], step: 1, category: 'Team', categories: ['Billing', 'Technical', 'Accounts'], targetCol: 'Ticket Goal', manager: 'support manager', lookupCol: 'Bonus Rate' },
  General: { title: 'Monthly Household Budget', role: 'Budget Planner', entity: 'expense category', entityPlural: 'expense categories', nameCol: 'Expense', names: ['Rent', 'Groceries', 'Electricity', 'Internet', 'Transport', 'Dining Out', 'Mobile', 'Insurance', 'Gym', 'Streaming', 'Clothing', 'Medical', 'Savings', 'Gifts', 'Books', 'Repairs', 'Water', 'Fuel'], periods: ['Jan', 'Feb', 'Mar', 'Apr'], measure: 'spending', type: 'currency', range: [40, 1800], step: 10, category: 'Type', categories: ['Needs', 'Wants', 'Savings'], targetCol: 'Budget', manager: 'family', lookupCol: 'Tax Rate', currency: '$' },
}

const ROWS: Record<ProblemLevel, number> = { beginner: 8, intermediate: 12, advanced: 16 }

/** Builds a themed problem that exercises exactly the requested concepts. */
export function builtinProblem(opts: { domain: Domain; level: ProblemLevel; concepts: ConceptId[]; seed: string; currency?: string }): ProblemDraft {
  const t = THEMES[opts.domain] ?? THEMES.General
  const R = rng(opts.seed)
  const n = ROWS[opts.level]
  const names = R.shuffle(t.names).slice(0, n)
  const cats = t.categories
  const want = new Set(opts.concepts)
  const needTarget = want.has('percent') || want.has('if')
  const needLookup = want.has('xlookup') || want.has('vlookup') || want.has('index-match')
  const currency = opts.currency ?? t.currency
  const round = (v: number) => Math.round(v / t.step) * t.step

  const columns: ProblemDraft['dataset']['columns'] = [{ name: t.nameCol, type: 'text' }, { name: t.category, type: 'text' }, ...t.periods.map((p) => ({ name: p, type: t.type }))]
  if (needTarget) columns.push({ name: t.targetCol, type: t.type })
  const rows = names.map((name, i) => {
    const base = R.int(t.range[0], t.range[1])
    const vals = t.periods.map(() => round(Math.max(t.range[0], Math.min(t.range[1], base + R.int(-0.25 * base, 0.25 * base)))))
    const row: (string | number)[] = [name, cats[i % cats.length], ...vals]
    if (needTarget) {
      const total = vals.reduce((a, b) => a + b, 0)
      // about two thirds of rows meet the target
      row.push(round(total * (R.next() < 0.65 ? 0.85 + R.next() * 0.1 : 1.05 + R.next() * 0.12)))
    }
    return row
  })

  const d: ProblemDraft = {
    title: t.title,
    context: `You work as a ${t.role}. Your ${t.manager} has shared the ${t.measure} of ${n} ${t.entityPlural} for ${t.periods[0]} to ${t.periods[t.periods.length - 1]}.`,
    objective: '',
    explanation: {
      situation: `Your ${t.manager} has given you ${t.measure} data for ${t.periods.length} periods.`,
      goal: '',
      data: `Each row is one ${t.entity}. The columns ${t.periods[0]} to ${t.periods[t.periods.length - 1]} hold the ${t.measure} for each period, and ${t.category} tells you which group they belong to.`,
      think: '',
    },
    concepts: [...want],
    estimatedMinutes: 10 + 5 * want.size,
    currencySymbol: currency,
    dataset: { columns, rows },
    rowTasks: [],
    summaryTasks: [],
    operations: [],
    instructions: [],
  }
  const first = t.periods[0]
  const last = t.periods[t.periods.length - 1]
  const goals: string[] = []
  const totalHeader = `Total ${cap(t.measure)}`
  const needTotal = ['sum', 'max', 'min', 'percent', 'if', 'sumif', 'averageif', 'sort', 'chart', 'formatting'].some((c) => want.has(c as ConceptId))
  if (needTotal) {
    d.rowTasks.push({ header: totalHeader, requirement: `Total ${t.measure} for each ${t.entity}`, concept: 'sum', formula: `=SUM([${first}]:[${last}])`, format: t.type })
    goals.push(`the total ${t.measure} of each ${t.entity}`)
  }
  if (want.has('average')) {
    d.rowTasks.push({ header: 'Average', requirement: `Average ${t.measure} per period for each ${t.entity}`, concept: 'average', formula: `=AVERAGE([${first}]:[${last}])`, format: t.type === 'currency' ? 'currency' : 'number' })
    goals.push(`the average ${t.measure} per period`)
  }
  if (want.has('round')) {
    d.rowTasks.push({ header: 'Average (1 dp)', requirement: `Average ${t.measure} rounded to 1 decimal place`, concept: 'round', formula: `=ROUND(AVERAGE([${first}]:[${last}]),1)`, format: 'number' })
    goals.push('a rounded average')
  }
  if (want.has('formula')) {
    d.rowTasks.push({ header: `Change ${first}–${last}`, requirement: `Change in ${t.measure} from ${first} to ${last} (${last} minus ${first})`, concept: 'formula', formula: `=[${last}]-[${first}]`, format: t.type })
    goals.push(`how much each ${t.entity} changed from ${first} to ${last}`)
  }
  if (want.has('percent')) {
    d.rowTasks.push({ header: 'Achieved %', requirement: `${totalHeader} as a percentage of ${t.targetCol}`, concept: 'percent', formula: `=[${totalHeader}]/[${t.targetCol}]`, format: 'percent' })
    goals.push(`what percentage of the ${t.targetCol.toLowerCase()} each ${t.entity} reached`)
  }
  if (want.has('if')) {
    d.rowTasks.push({ header: 'Status', requirement: `Show "Met" when ${totalHeader} is at least ${t.targetCol}, otherwise "Not met"`, concept: 'if', formula: `=IF([${totalHeader}]>=[${t.targetCol}],"Met","Not met")` })
    goals.push(`whether each ${t.entity} met the ${t.targetCol.toLowerCase()}`)
  }
  if (want.has('concat')) {
    d.rowTasks.push({ header: 'Label', requirement: `Join the ${t.nameCol} and ${t.category} as “Name (${t.category})”`, concept: 'concat', formula: `=CONCAT([${t.nameCol}]," (",[${t.category}],")")` })
    goals.push('a readable label for each row')
  }
  if (needLookup) {
    const rates = cats.map((_, i) => +(0.02 + 0.01 * ((i * 3 + R.int(0, 2)) % 6)).toFixed(2))
    d.lookup = { name: 'Lookup', columns: [{ name: t.category, type: 'text' }, { name: t.lookupCol, type: 'percent' }], rows: cats.map((c, i) => [c, rates[i]]) }
    const fn = want.has('xlookup') ? 'xlookup' : want.has('vlookup') ? 'vlookup' : 'index-match'
    const formula =
      fn === 'xlookup'
        ? `=XLOOKUP([${t.category}],[Lookup.${t.category}:all],[Lookup.${t.lookupCol}:all])`
        : fn === 'vlookup'
          ? `=VLOOKUP([${t.category}],[Lookup.${t.category}:all]:[Lookup.${t.lookupCol}:all],2,FALSE)`
          : `=INDEX([Lookup.${t.lookupCol}:all],MATCH([${t.category}],[Lookup.${t.category}:all],0))`
    d.rowTasks.push({ header: t.lookupCol, requirement: `Look up the ${t.lookupCol.toLowerCase()} for each row's ${t.category} from the Lookup sheet`, concept: fn, formula, format: 'percent' })
    goals.push(`the ${t.lookupCol.toLowerCase()} that applies to each ${t.entity}`)
  }
  // summaries
  const sumCol = needTotal ? totalHeader : first
  const catPick = cats[R.int(0, cats.length - 1)]
  if (want.has('sum') && needTotal) d.summaryTasks.push({ label: 'Grand total', requirement: `Grand total of ${totalHeader}`, concept: 'sum', formula: `=SUM([${totalHeader}:all])`, format: t.type, underColumn: totalHeader })
  if (want.has('average')) d.summaryTasks.push({ label: 'Average total', requirement: `Average of ${sumCol} across all ${t.entityPlural}`, concept: 'average', formula: `=AVERAGE([${sumCol}:all])`, format: t.type === 'currency' ? 'currency' : 'number', underColumn: sumCol })
  if (want.has('max')) {
    d.summaryTasks.push({ label: 'Highest', requirement: `Highest ${sumCol}`, concept: 'max', formula: `=MAX([${sumCol}:all])`, format: t.type, underColumn: sumCol })
    goals.push(`the highest-performing ${t.entity}`)
  }
  if (want.has('min')) d.summaryTasks.push({ label: 'Lowest', requirement: `Lowest ${sumCol}`, concept: 'min', formula: `=MIN([${sumCol}:all])`, format: t.type, underColumn: sumCol })
  if (want.has('count')) d.summaryTasks.push({ label: `Number of ${t.entityPlural}`, requirement: `How many ${t.entityPlural} are listed`, concept: 'count', formula: `=COUNTA([${t.nameCol}:all])`, format: 'integer', underColumn: sumCol })
  if (want.has('countif')) d.summaryTasks.push({ label: `${catPick} count`, requirement: `How many ${t.entityPlural} are in ${t.category} “${catPick}”`, concept: 'countif', formula: `=COUNTIF([${t.category}:all],"${catPick}")`, format: 'integer', underColumn: sumCol })
  if (want.has('sumif')) d.summaryTasks.push({ label: `${catPick} total`, requirement: `Total ${sumCol} for ${t.category} “${catPick}”`, concept: 'sumif', formula: `=SUMIF([${t.category}:all],"${catPick}",[${sumCol}:all])`, format: t.type, underColumn: sumCol })
  if (want.has('averageif')) d.summaryTasks.push({ label: `${catPick} average`, requirement: `Average ${sumCol} for ${t.category} “${catPick}”`, concept: 'averageif', formula: `=AVERAGEIF([${t.category}:all],"${catPick}",[${sumCol}:all])`, format: t.type === 'currency' ? 'currency' : 'number', underColumn: sumCol })
  if (want.has('sort')) {
    d.operations.push({ type: 'sort', column: sumCol, order: 'desc', requirement: `Sort the table so the highest ${sumCol} is at the top` })
    goals.push('a ranking from best to worst')
  }
  if (want.has('chart')) d.operations.push({ type: 'chart', chartType: 'column', labelColumn: t.nameCol, valueColumn: sumCol, requirement: `Insert a column chart of ${sumCol} by ${t.nameCol}` })
  if (want.has('formatting')) d.operations.push({ type: 'format', column: sumCol, format: t.type === 'currency' ? 'currency' : 'percent', requirement: `Format the ${sumCol} column as ${t.type === 'currency' ? 'Currency' : 'Percent'}` })

  d.concepts = [...new Set<ConceptId>([...want, ...d.rowTasks.map((x) => x.concept), ...d.summaryTasks.map((x) => x.concept)])]
  d.objective = `Work out ${listJoin(goals.length ? goals : [`the key figures for each ${t.entity}`])}.`
  d.explanation.goal = `They want to know ${listJoin(goals.slice(0, 3)) || 'the key figures'}.`
  d.explanation.think = needTotal ? `What do you think we need to do with the ${t.periods.length} ${t.measure} values in each row?` : `Which columns do you think you will need for this?`
  d.instructions = [
    ...d.rowTasks.map((x) => `${x.requirement} (column “${x.header}”).`),
    ...d.summaryTasks.map((x) => `${x.requirement} (row “${x.label}”).`),
    ...d.operations.map((o) => `${o.requirement}.`),
  ]
  return d
}

const cap = (s: string) => s.replace(/\b\w/g, (m) => m.toUpperCase())
const listJoin = (a: string[]) => (a.length <= 1 ? a.join('') : `${a.slice(0, -1).join(', ')} and ${a[a.length - 1]}`)
