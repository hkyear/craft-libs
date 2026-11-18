//! Tables: the table sprms (sgc 5) carried by a row-end (TTP) mark's grpprl, and the
//! assembly of cell-ended paragraphs into [`Table`] blocks ([MS-DOC] §2.6.3, §2.4.3).

use std::sync::Arc;

use craft_wordmodel::Rgb;
use craft_wordmodel::para::Paragraph;
use craft_wordmodel::props::{Align, Borders, CellProps, HeightRule, RowProps, TableProps, VAlign, VMerge};
use craft_wordmodel::table::{Cell, Row, Table};
use craft_wordmodel::{Block, Blocks, para_block};

use crate::fmt;
use crate::sprm::{self, Prl};

// Paragraph sprms for table membership.
const P_IN_TABLE: u16 = 0x2416;
const P_TTP: u16 = 0x2417;
/// `sprmPFInnerTableCell` / `sprmPFInnerTtp`: a paragraph mark ending a cell, or a row, of a
/// table nested in another's cell ([MS-DOC] §2.6.2); its mark is 0x0D, not 0x07.
const P_INNER_CELL: u16 = 0x244B;
const P_INNER_TTP: u16 = 0x244C;
/// `sprmPItap`: table depth (1 = an outermost table, 2 = one in its cell, …).
const P_ITAP: u16 = 0x6649;
/// `sprmPDtap`: signed adjustment to the depth set by earlier depth sprms.
const P_DTAP: u16 = 0x664A;
/// Deepest table nesting we build, as the .docx reader's; deeper rows stay paragraphs in the
/// cell of the deepest table built.
const MAX_TABLE_DEPTH: u32 = 24;

// Table sprms ([MS-DOC] §2.6.3).
const T_JC: u16 = 0x5400;
const T_JC_90: u16 = 0x548A;
const T_DXA_LEFT: u16 = 0x9601;
const T_DXA_GAP_HALF: u16 = 0x9602;
const T_TABLE_HEADER: u16 = 0x3404;
const T_TABLE_BORDERS_80: u16 = 0xD605;
const T_ROW_HEIGHT: u16 = 0x9407;
const T_DEF_TABLE: u16 = 0xD608;
const T_DEF_TABLE_SHD_80: u16 = 0xD609;
/// `sprmTDefTableShd` / `Shd2nd` / `Shd3rd`: 10-byte `Shd`s for cells 0-21, 22-43, 44-62.
const T_DEF_TABLE_SHD: u16 = 0xD612;
const T_DEF_TABLE_SHD_2ND: u16 = 0xD616;
const T_DEF_TABLE_SHD_3RD: u16 = 0xD60C;
/// `sprmTDefTableShdRaw` / `Raw2nd` / `Raw3rd`: the same `Shd` arrays over the same cell
/// ranges, which Word 2007+ writes next to the above.
const T_DEF_TABLE_SHD_RAW: u16 = 0xD670;
const T_DEF_TABLE_SHD_RAW_2ND: u16 = 0xD671;
const T_DEF_TABLE_SHD_RAW_3RD: u16 = 0xD672;
const T_CANT_SPLIT_90: u16 = 0x3403;
const T_MERGE: u16 = 0x5624;
const T_VERT_MERGE: u16 = 0xD62B;
const T_VERT_ALIGN: u16 = 0xD62C;
const T_SET_SHD: u16 = 0xD62D;
const T_CANT_SPLIT: u16 = 0x3466;
const T_CELL_WIDTH: u16 = 0xD635;

/// Most grid columns we build from a `sprmTDefTable` (the model's own cap is 63; one more
/// edge is fine but merged rows can sum spans slightly over, so clamp when assembling).
const MAX_COLS: usize = 63;

/// What a paragraph mark says about its table context, from its PAPX.
#[derive(Default)]
pub(crate) struct RowInfo {
    pub(crate) in_table: bool,
    pub(crate) ttp: bool,
    /// The mark itself was 0x07 (a cell end), set by the walker.
    pub(crate) cell_mark: bool,
    /// Table depth from `sprmPItap` / `sprmPDtap` (0 when absent, read as 1 for a table mark).
    pub(crate) depth: u32,
    /// `sprmPFInnerTableCell`: the mark ends a cell of a nested table.
    inner_cell: bool,
    /// `sprmPFInnerTtp`: the mark ends a row of a nested table.
    inner_ttp: bool,
    /// Cell-edge positions from `sprmTDefTable`, in the file's native i16 twips — [MS-DOC]
    /// defines them per row, so each row keeps its own; ascending in well-formed rows, the
    /// first edge usually 0 (the row origin). Rows without a `sprmTDefTable` keep this empty.
    /// Conversion to points happens once, at the model boundary.
    pub(crate) edges: Vec<i16>,
    height: Option<f32>,
    height_rule: HeightRule,
    header: bool,
    cant_split: bool,
    align: Option<Align>,
    indent: Option<f32>,
    borders: Option<Borders>,
    /// Horizontal merges: (itcFirst, itcLim) cell-index ranges.
    merges: Vec<(u8, u8)>,
    /// Per-cell overrides keyed by (pre-merge) cell index.
    vmerge: Vec<(u8, VMerge)>,
    /// Later entries win: the `Shd` arrays override the `Shd80` fallback before them, and an
    /// automatic `Shd` (`None`) clears it.
    shading: Vec<(u8, Option<Rgb>)>,
    /// Native sprm units: twips (ftsWidth 3) and 50ths of a percent (ftsWidth 2).
    widths: Vec<(u8, i16)>,
    width_pcts: Vec<(u8, u16)>,
    valign: Vec<(u8, VAlign)>,
}

/// Read the table context out of a paragraph mark's PAPX grpprl.
impl RowInfo {
    /// Does the mark belong to a table: inside one, ending a row, or ending a cell?
    pub(crate) fn is_table(&self) -> bool {
        self.in_table || self.ttp || self.cell_mark
    }

    /// The table depth the mark is at: `sprmPItap`, or 1 when a table mark states none.
    fn level(&self) -> u32 {
        self.depth.max(1)
    }

    /// Does the mark end a cell, and a row, of the table at `level`? Outermost cells end with
    /// 0x07 and rows with `sprmPFTtp`; nested ones with the inner sprms on a 0x0D mark.
    fn ends_at(&self, level: u32) -> (bool, bool) {
        if level <= 1 { (self.cell_mark, self.ttp) } else { (self.inner_cell || self.inner_ttp, self.inner_ttp) }
    }

    /// The row height from `sprmTDyaRowHeight`, for tests.
    #[cfg(test)]
    pub(crate) fn height(&self) -> Option<f32> {
        self.height
    }

    /// The fill cell `i` ends up with (the last shading entry for it wins), for tests.
    #[cfg(test)]
    pub(crate) fn shading_of(&self, i: u8) -> Option<Rgb> {
        self.shading.iter().rfind(|(k, _)| *k == i).and_then(|(_, s)| *s)
    }

    /// The width cell `i` ends up with, for tests.
    #[cfg(test)]
    pub(crate) fn width_of(&self, i: u8) -> Option<f32> {
        self.widths.iter().rfind(|(k, _)| *k == i).map(|(_, w)| *w as f32 / 20.0)
    }

    /// How many per-cell overrides the row keeps, for tests.
    #[cfg(test)]
    pub(crate) fn overrides(&self) -> usize {
        self.vmerge.len() + self.shading.len() + self.widths.len() + self.width_pcts.len() + self.valign.len() + self.merges.len()
    }

    /// The width percentage cell `i` ends up with, for tests.
    #[cfg(test)]
    pub(crate) fn width_pct_of(&self, i: u8) -> Option<f32> {
        self.width_pcts.iter().rfind(|(k, _)| *k == i).map(|(_, w)| *w as f32 / 50.0)
    }
}

/// Set cell `c`'s override to `v`, replacing an earlier one: one slot per cell, however often
/// a PAPX repeats the sprm, so a shared PAPX can't pile up copies on every paragraph mark.
fn set<T>(slots: &mut Vec<(u8, T)>, c: u8, v: T) {
    match slots.iter_mut().find(|(k, _)| *k == c) {
        Some(slot) => slot.1 = v,
        None => slots.push((c, v)),
    }
}

pub(crate) fn decode(papx: &[u8]) -> RowInfo {
    let mut r = RowInfo::default();
    let mut gap_half: Option<i16> = None;
    let mut dxa_left: Option<i16> = None;
    for prl in sprm::iter(papx) {
        // The membership sprms are paragraph sprms (sgc 1); the rest here are table sprms
        // (sgc 5), so match opcodes directly.
        match prl.op {
            P_IN_TABLE => r.in_table = prl.operand.first().copied().unwrap_or(0) != 0,
            P_TTP => r.ttp = prl.operand.first().copied().unwrap_or(0) != 0,
            P_INNER_CELL => r.inner_cell = prl.operand.first().copied().unwrap_or(0) != 0,
            P_INNER_TTP => r.inner_ttp = prl.operand.first().copied().unwrap_or(0) != 0,
            P_ITAP => r.depth = i32_of(&prl).unwrap_or(0).max(0) as u32,
            // Apply in sprm order; malformed adjustments must not underflow or overflow.
            P_DTAP => r.depth = r.depth.saturating_add_signed(i32_of(&prl).unwrap_or(0)),
            T_DEF_TABLE => r.edges = def_table_edges(prl.operand),
            T_TABLE_HEADER => r.header = prl.operand.first().copied().unwrap_or(0) != 0,
            T_CANT_SPLIT | T_CANT_SPLIT_90 => r.cant_split = prl.operand.first().copied().unwrap_or(0) != 0,
            T_ROW_HEIGHT => {
                if let Some(v) = s16(&prl) {
                    if v < 0 {
                        r.height_rule = HeightRule::Exact;
                        r.height = Some((-(v as f32) / 20.0).clamp(0.0, 1584.0));
                    } else {
                        r.height_rule = HeightRule::AtLeast;
                        r.height = Some((v as f32 / 20.0).clamp(0.0, 1584.0));
                    }
                }
            }
            T_JC | T_JC_90 => {
                r.align = match u16_of(&prl).unwrap_or(0) {
                    1 => Some(Align::Center),
                    2 => Some(Align::Right),
                    _ => Some(Align::Left),
                };
            }
            T_DXA_LEFT => dxa_left = s16(&prl),
            T_DXA_GAP_HALF => gap_half = s16(&prl),
            T_TABLE_BORDERS_80 => r.borders = table_borders(prl.operand),
            T_MERGE => {
                // Rows apply only the first MAX_COLS merges.
                if let [a, b, ..] = prl.operand
                    && b > a
                    && r.merges.len() < MAX_COLS
                {
                    r.merges.push((*a, *b));
                }
            }
            T_VERT_MERGE => {
                // cb, itc, flag: 1 = continues the merge from above, 3 = restarts it.
                if let [_, itc, flag, ..] = prl.operand {
                    let m = match flag {
                        3 => VMerge::Restart,
                        1 => VMerge::Continue,
                        _ => VMerge::None,
                    };
                    // The first override for a cell wins.
                    if m != VMerge::None && !r.vmerge.iter().any(|(k, _)| k == itc) {
                        r.vmerge.push((*itc, m));
                    }
                }
            }
            T_SET_SHD => {
                if let Some(rgb) = prl.operand.get(3..).and_then(shd_rgb)
                    && let [a, b, ..] = prl.operand.get(1..3).unwrap_or(&[])
                {
                    for c in *a..(*b).min(MAX_COLS as u8) {
                        set(&mut r.shading, c, Some(rgb));
                    }
                }
            }
            T_DEF_TABLE_SHD_80 => {
                let cells = prl.operand.get(1..).unwrap_or(&[]);
                for (i, b) in cells.as_chunks::<2>().0.iter().take(MAX_COLS).enumerate() {
                    if let Some(rgb) = shd80_rgb(b) {
                        set(&mut r.shading, i as u8, Some(rgb));
                    }
                }
            }
            T_DEF_TABLE_SHD | T_DEF_TABLE_SHD_RAW | T_DEF_TABLE_SHD_2ND | T_DEF_TABLE_SHD_RAW_2ND | T_DEF_TABLE_SHD_3RD | T_DEF_TABLE_SHD_RAW_3RD => {
                // Each array covers its own cell range (0-21, 22-43, 44-62); extra entries
                // are ignored so an over-long array can't override the next range.
                let (first, lim) = match prl.op {
                    T_DEF_TABLE_SHD_2ND | T_DEF_TABLE_SHD_RAW_2ND => (22, 44),
                    T_DEF_TABLE_SHD_3RD | T_DEF_TABLE_SHD_RAW_3RD => (44, MAX_COLS),
                    _ => (0, 22),
                };
                let cells = prl.operand.get(1..).unwrap_or(&[]);
                let raw = matches!(prl.op, T_DEF_TABLE_SHD_RAW | T_DEF_TABLE_SHD_RAW_2ND | T_DEF_TABLE_SHD_RAW_3RD);
                for (i, b) in cells.as_chunks::<10>().0.iter().take(lim - first).enumerate() {
                    // In the Raw arrays a nil Shd (ipat 0xFFFF) means "as the table style
                    // shades it", so it leaves any earlier fill alone.
                    if raw && b.get(8..10) == Some(&[0xFF, 0xFF]) {
                        continue;
                    }
                    set(&mut r.shading, (first + i) as u8, shd_rgb(b));
                }
            }
            T_CELL_WIDTH => {
                // Operand: cb (1 byte), itcFirst (1 byte), itcLim (1 byte), ftsWidth (1 byte), wWidth (2 bytes).
                // [MS-DOC] §2.9.290 sprmTCellWidth (0xD635).
                // ftsWidth: 3 = twips (dxa), 2 = percentage (50ths of a percent).
                if let [_, a, b, fts, w0, w1, ..] = prl.operand {
                    if *fts == 3 {
                        // 1584 pt = 31680 twips, the clamp the old f32 path applied.
                        let w = i16::from_le_bytes([*w0, *w1]).clamp(0, 31680);
                        for c in *a..(*b).min(MAX_COLS as u8) {
                            set(&mut r.widths, c, w);
                        }
                    } else if *fts == 2 {
                        let pct = u16::from_le_bytes([*w0, *w1]).min(5000);
                        for c in *a..(*b).min(MAX_COLS as u8) {
                            set(&mut r.width_pcts, c, pct);
                        }
                    }
                }
            }
            T_VERT_ALIGN => {
                if let [_, a, b, va, ..] = prl.operand {
                    let v = match va {
                        1 => VAlign::Center,
                        2 => VAlign::Bottom,
                        _ => VAlign::Top,
                    };
                    for c in *a..(*b).min(MAX_COLS as u8) {
                        set(&mut r.valign, c, v);
                    }
                }
            }
            _ => {}
        }
    }
    // The table indent is the row origin minus half the cell gap.
    if let Some(l) = dxa_left {
        let gap = gap_half.unwrap_or(0) as f32 / 20.0;
        r.indent = Some(((l as f32 / 20.0) - gap).clamp(-1584.0, 1584.0));
    }
    r
}

/// A finished paragraph plus the bookkeeping the block assembly needs.
pub(crate) struct ParaOut {
    pub(crate) para: Paragraph,
    /// CP just past the terminating mark.
    pub(crate) end_cp: u32,
    /// Table context from the mark's PAPX; `None` for a paragraph outside a table.
    pub(crate) row: Option<Box<RowInfo>>,
}

/// Group consecutive in-table paragraphs into [`Table`] blocks; the rest pass through. A
/// table nested in a cell (`sprmPItap` above the table's own depth) becomes a table inside
/// that cell, to [`MAX_TABLE_DEPTH`]; deeper rows stay paragraphs so no text is lost.
pub(crate) fn assemble(paras: Vec<ParaOut>) -> Blocks {
    assemble_at(&mut paras.into_iter().peekable(), 1)
}

/// Assemble one depth from a shared consuming cursor. A nested call stops before the first
/// shallower mark, leaving it for its parent. Keeping one input allocation avoids retaining
/// a near-full paragraph buffer per depth on large, deeply nested documents.
fn assemble_at(paras: &mut std::iter::Peekable<impl Iterator<Item = ParaOut>>, level: u32) -> Blocks {
    let mut out: Blocks = Vec::new();
    // Rows are finalized in two phases: cell marks collect raw cells plus the row's own
    // `sprmTDefTable` edges; flush_table derives one table-wide grid (the union of every
    // row's edges, as Word does when it saves) and maps each row onto it.
    let mut rows_raw: Vec<(Vec<Cell>, Box<RowInfo>, bool)> = Vec::new();
    let mut cells: Vec<Cell> = Vec::new();
    let mut cell: Blocks = Vec::new();
    let mut props = TableProps::default();
    let mut open = false;

    fn flush_cell(cells: &mut Vec<Cell>, cell: &mut Blocks) {
        let mut blocks = std::mem::take(cell);
        if blocks.is_empty() {
            blocks.push(para_block(Paragraph::new()));
        }
        cells.push(Cell { props: CellProps { span: 1, ..Default::default() }, blocks, controls: Default::default() });
    }
    /// Map one row's cells onto the unified grid from its own TDefTable edges: a cell
    /// spans every grid column whose left edge falls inside it, zero-width edges are
    /// shadows folded into the previous cell, horizontal merges span the sum of the cells
    /// they join, and a row narrower than the table gains empty leading/trailing cells
    /// (LibreOffice's empty edge columns) so every row spans the whole grid.
    fn finish_row(cells: Vec<Cell>, info: &RowInfo, unified: &[i16]) -> Row {
        let mut cells = cells;
        // Per-cell overrides first (sprm cell indices count pre-merge cells); merges map
        // below, after spans, so they span the sum of the cells they join.
        for (i, c) in cells.iter_mut().enumerate() {
            let i = i as u8;
            if let Some((_, m)) = info.vmerge.iter().find(|(k, _)| *k == i) {
                c.props.vmerge = *m;
            }
            if let Some((_, w)) = info.widths.iter().rfind(|(k, _)| *k == i) {
                c.props.width = Some(*w as f32 / 20.0);
            }
            if let Some((_, pct)) = info.width_pcts.iter().rfind(|(k, _)| *k == i) {
                c.props.width_pct = Some(*pct as f32 / 50.0);
            }
            if let Some((_, s)) = info.shading.iter().rfind(|(k, _)| *k == i) {
                c.props.shading = *s;
            }
            if let Some((_, v)) = info.valign.iter().rfind(|(k, _)| *k == i) {
                c.props.valign = Some(*v);
            }
        }
        let edges = &info.edges;
        if edges.len() >= 2 && cells.len() + 1 == edges.len() {
            // Spans first, in row-column space; merges below sum them.
            // Exact integer comparison: edges and the unified grid share the file's
            // twip unit, so no epsilon is needed at all.
            let span_of = |i: usize| -> u32 {
                let (a, b) = (edges[i], edges[i + 1]);
                unified.iter().filter(|&&g| g >= a && g < b).count() as u32
            };
            for i in 0..cells.len() {
                let span = span_of(i);
                if span == 0 && i > 0 {
                    // A zero-width edge pair is a shadow cell: fold it into the previous one.
                    let (left, right) = cells.split_at_mut(i);
                    let prev = left.last_mut().expect("i > 0");
                    prev.blocks.extend(right[0].blocks.drain(..).filter(|x| x.as_para().is_none_or(|p| !p.is_empty())));
                    right[0].props.span = 0;
                } else {
                    cells[i].props.span = span.clamp(1, 63);
                }
            }
            cells.retain(|c| c.props.span > 0);
            // The row's leading edge may sit right of the table's (an indented row): an
            // empty cell covers the columns before it, another covers the ones after it.
            let mut spans: u32 = cells.iter().map(|c| c.props.span).sum();
            let grid_cols = unified.len().saturating_sub(1) as u32;
            if grid_cols > 0 {
                if edges[0] > unified[0] {
                    let lead = unified.iter().filter(|&&g| g < edges[0]).count() as u32;
                    if lead > 0 {
                        spans += lead;
                        cells.insert(0, Cell { props: CellProps { span: lead.min(63), ..Default::default() }, blocks: vec![para_block(Paragraph::new())], controls: Default::default() });
                    }
                }
                if spans < grid_cols {
                    cells.push(Cell { props: CellProps { span: (grid_cols - spans).min(63), ..Default::default() }, blocks: vec![para_block(Paragraph::new())], controls: Default::default() });
                }
            }
        }
        if std::env::var_os("CRAFT_DOCBIN_DEBUG").is_some() {
            eprintln!("[finish] edges={edges:?} cells={} spans={:?} merges={:?} unified={}cols", cells.len(), cells.iter().map(|c| c.props.span).collect::<Vec<_>>(), info.merges, unified.len().saturating_sub(1));
        }
        // Horizontal merges: (itcFirst, itcLim) ranges in pre-merge cell indices; the
        // merged cell spans the sum of the cells it joins.
        let mut cells = cells;
        for &(a, b) in info.merges.iter().take(MAX_COLS) {
            let a = a as usize;
            let b = (b as usize).min(cells.len());
            if a < b && b <= cells.len() && b - a > 1 {
                let span: u32 = cells[a..b].iter().map(|c| c.props.span).sum();
                let removed: Vec<Cell> = cells.drain(a + 1..b).collect();
                if let Some(first) = cells.get_mut(a) {
                    first.props.span = span.clamp(1, 63);
                    for c in removed {
                        first.blocks.extend(c.blocks.into_iter().filter(|x| x.as_para().is_none_or(|p| !p.is_empty())));
                    }
                    if first.blocks.len() > 1 && first.blocks.first().and_then(|x| x.as_para()).is_some_and(Paragraph::is_empty) {
                        first.blocks.remove(0);
                    }
                }
            }
        }
        Row {
            props: RowProps {
                height: info.height,
                height_rule: info.height_rule,
                header: info.header,
                cant_split: info.cant_split,
                ..Default::default()
            },
            cells,
            controls: Default::default(),
        }
    }

    fn flush_table(out: &mut Blocks, rows_raw: &mut Vec<(Vec<Cell>, Box<RowInfo>, bool)>, cells: &mut Vec<Cell>, cell: &mut Blocks, props: &TableProps) {
        if !cells.is_empty() || !cell.is_empty() {
            flush_cell(cells, cell);
            // A stray trailing row (no row end of its own) keeps today's shape: one span-1
            // cell per column of the grid, sized by the grid where it has one.
            rows_raw.push((std::mem::take(cells), Box::new(RowInfo::default()), true));
        }
        if rows_raw.is_empty() {
            return;
        }
        // One grid for the whole table, as Word saves one: the union of every row's
        // `sprmTDefTable` edges, sorted and deduplicated. Rows may redefine the grid
        // per row (LibreOffice's bands); the union lets each row's own segmentation map
        // onto shared columns. Rows without any edges contribute nothing.
        let mut unified: Vec<i16> = rows_raw.iter().filter(|(_, i, _)| i.edges.len() >= 2).flat_map(|(_, i, _)| i.edges.iter().copied()).collect();
        unified.sort_unstable();
        unified.dedup();
        if unified.len() < 2 || unified.len() > MAX_COLS + 1 {
            // Degenerate or over-wide unions fall back to the first row that has edges.
            unified = rows_raw.iter().find(|(_, i, _)| i.edges.len() >= 2).map_or(Vec::new(), |(_, i, _)| {
                let mut e: Vec<i16> = i.edges.iter().copied().collect();
                e.sort_unstable();
                e.dedup();
                e
            });
        }
        // One conversion to points at the model boundary; i16 differences may exceed
        // i16's range, so widen for the subtraction.
        let grid: Vec<f32> = unified
            .windows(2)
            .map(|w| ((w[1] as i32 - w[0] as i32).max(0)) as f32 / 20.0)
            .take(MAX_COLS)
            .collect();
        let mut rows: Vec<Row> = Vec::with_capacity(rows_raw.len());
        for (cells, info, stray) in rows_raw.drain(..) {
            let mut row = finish_row(cells, &info, &unified);
            // A stray trailing row (no row end of its own): size its cells from the grid.
            if stray {
                for (i, c) in row.cells.iter_mut().enumerate() {
                    if c.props.width.is_none() && c.props.span == 1 {
                        c.props.width = grid.get(i).copied();
                    }
                }
            }
            rows.push(row);
        }
        // A table whose rows carry no width intent (no `sprmTCellWidth`, no percentage
        // widths) keeps its widths only in the `sprmTDefTable` boundary array. Word shows
        // those boundaries; without a layout hint the layout engine content-fits every
        // column, and boundary-only tables (government and exam forms) collapse. Declare
        // the table fixed at its saved total so the grid stands as written.
        let mut props = props.clone();
        if !grid.is_empty() && !rows.iter().any(|r| r.cells.iter().any(|c| c.props.width.is_some() || c.props.width_pct.is_some())) {
            props.fixed = true;
            props.width = Some(grid.iter().sum());
        }
        out.push(Arc::new(Block::Table(Table {
            props,
            grid,
            rows,
            controls: Default::default(),
        })));
    }

    let dbg = std::env::var_os("CRAFT_DOCBIN_DEBUG").is_some();
    while let Some(next) = paras.peek() {
        let next_level = next.row.as_ref().map_or(0, |r| r.level());
        if dbg && next.row.is_some() {
            let r = next.row.as_ref().unwrap();
            eprintln!(
                "[mark] lvl={} in_table={} ttp={} cell_mark={} inner_cell={} inner_ttp={} edges={} depth={} text={:?}",
                r.level(), r.in_table, r.ttp, r.cell_mark, r.inner_cell, r.inner_ttp, r.edges.len(), r.depth, next.para.text.as_str().chars().take(12).collect::<String>()
            );
        }
        if level > 1 && next_level < level {
            break;
        }
        if next_level > level && level < MAX_TABLE_DEPTH {
            open = true;
            cell.extend(assemble_at(paras, level + 1));
            continue;
        }
        let Some(mut p) = paras.next() else { break };
        if next_level > level {
            // Beyond the depth cap, preserve the text in the deepest cell we build.
            open = true;
            cell.push(para_block(p.para));
            continue;
        }
        // A cell mark with no table context at all: Word still ends the open row with it
        // (Aspose-written tables terminate rows with a naked 0x07 whose PAPX carries no
        // table sprm), and a naked 0x07 outside any table is a plain paragraph, not the
        // one-cell table today's code would invent from it.
        if p.row.as_ref().is_some_and(|r| {
            r.cell_mark && !r.in_table && !r.ttp && !r.inner_cell && !r.inner_ttp && r.depth == 0 && p.para.text.is_empty()
        }) {
            if !open {
                out.push(para_block(p.para));
                paras.next();
                continue;
            }
            if !cells.is_empty() || !cell.is_empty() {
                if !cell.is_empty() {
                    flush_cell(&mut cells, &mut cell);
                }
                rows_raw.push((std::mem::take(&mut cells), p.row.take().unwrap_or_default(), false));
            }
            paras.next();
            continue;
        }
        let Some(info) = p.row else {
            if open {
                flush_table(&mut out, &mut rows_raw, &mut cells, &mut cell, &props);
                // Table props are per-table (first row carrying the sprm wins within the
                // table): reset the accumulator so one table's borders, align, indent and
                // cell margins never leak into the next (the repro's first table wrote
                // Word's default 108-twip gap, and every later table — including the
                // change history with its own 40-twip gap — inherited 5.4 pt margins).
                props = TableProps::default();
                open = false;
            }
            out.push(para_block(p.para));
            continue;
        };
        open = true;
        if info.borders.is_some() && props.borders.is_none() {
            props.borders = info.borders;
        }
        if props.align.is_none() {
            props.align = info.align;
        }
        if props.indent.is_none() {
            props.indent = info.indent;
        }
        let (cell_end, row_end) = info.ends_at(level);
        if cell_end {
            // A row-end (TTP) mark terminates the row. Word writes rows as N content
            // cells each ended by their own cell mark, followed by an empty row-end mark
            // that carries the row's TDefTable (the 3GPP change history: 8 content marks
            // + an empty TTP whose TDefTable defines itcMac = 8). When that mark ends no
            // content of its own and the TDefTable already has a column per flushed cell,
            // it is a pure row terminator, not the end of a phantom (N+1)th cell — such a
            // cell has no column (width None) and shifts every real cell's span one grid
            // column over. A TTP mark with content, or one short of the TDefTable's
            // columns, still ends a real cell (rows without a TDefTable keep today's
            // shape: the count is unknowable, so the mark always flushes).
            let row_end_only = row_end && p.para.text.is_empty() && cell.is_empty() && !info.edges.is_empty() && cells.len() + 1 >= info.edges.len();
            // ...but a row-end mark with no cells before it still makes a (one-cell) row.
            let row_end_only = row_end_only && !cells.is_empty();
            if !p.para.text.is_empty() {
                cell.push(para_block(p.para));
            }
            if !row_end_only {
                flush_cell(&mut cells, &mut cell);
            }
            if row_end && (!cells.is_empty() || !cell.is_empty()) {
                rows_raw.push((std::mem::take(&mut cells), info, false));
            }
        } else {
            cell.push(para_block(p.para));
        }
    }
    if open {
        flush_table(&mut out, &mut rows_raw, &mut cells, &mut cell, &props);
    }
    out
}

/// Cell-edge positions (points) from a `sprmTDefTable` operand: cb u16, itcMac u8, then
/// itcMac+1 twip boundaries followed by the TC80s (skipped). Edges are returned as stored:
/// per-row, in operand order; sorting and capping happen when the table grid is assembled.
/// Raw twip edges; the /20 to points happens once, when the model grid is built.
fn def_table_edges(operand: &[u8]) -> Vec<i16> {
    let itc = operand.get(2).copied().unwrap_or(0) as usize;
    if itc == 0 {
        return Vec::new();
    }
    // itcMac is a u8, so this allocation stays tiny; the loop stops at the operand's end.
    let mut edges = Vec::with_capacity((itc + 1).min(256));
    for i in 0..=itc {
        match operand.get(3 + i * 2..3 + i * 2 + 2) {
            Some(b) => edges.push(i16::from_le_bytes([b[0], b[1]])),
            None => break,
        }
    }
    edges
}

/// Table borders from `sprmTTableBorders80`: cb (0x18) then six Brc80s in the order top,
/// left, bottom, right, inside-horizontal, inside-vertical, read as paragraph borders are.
fn table_borders(operand: &[u8]) -> Option<Borders> {
    let b = |i: usize| operand.get(1 + i * 4..1 + i * 4 + 4).and_then(fmt::brc80);
    Some(Borders { top: b(0), left: b(1), bottom: b(2), right: b(3), between: b(4), inside_v: b(5) })
}

/// Background colour from a 10-byte `Shd`: cvFore, cvBack (COLORREFs), ipat. Solid patterns
/// show the foreground; clear (0) and the other patterns show the background (Word writes a
/// plain cell fill as clear with cvBack); nil (0xFFFF) and automatic colours are no fill.
fn shd_rgb(b: &[u8]) -> Option<Rgb> {
    let cv = |b: &[u8]| -> Option<Rgb> {
        match b {
            [_, _, _, 0xFF] => None, // fAuto
            [r, g, bl, _] => Some(Rgb(*r, *g, *bl)),
            _ => None,
        }
    };
    let ipat = b.get(8..10).map(|x| u16::from_le_bytes([x[0], x[1]]))?;
    match ipat {
        0xFFFF => None,
        1 => cv(b.get(0..4)?),
        _ => cv(b.get(4..8)?),
    }
}

/// Background colour from a 2-byte `Shd80` (icoFore, icoBack, ipat bit-fields). As in
/// [`shd_rgb`], solid shows the foreground and clear (0) the background, which is how Word
/// 97 writes a plain fill; ico 0 is automatic (no fill).
fn shd80_rgb(b: &[u8]) -> Option<Rgb> {
    let v = u16::from_le_bytes([*b.first()?, *b.get(1)?]);
    let fore = v & 0x1F;
    let back = (v >> 5) & 0x1F;
    let ipat = v >> 10;
    if fore == 0x1F && back == 0x1F {
        return None;
    }
    fmt::ico(if ipat == 1 { fore as u8 } else { back as u8 })
}

fn u16_of(p: &Prl) -> Option<u16> {
    match p.operand {
        [b0, b1, ..] => Some(u16::from_le_bytes([*b0, *b1])),
        _ => None,
    }
}

fn s16(p: &Prl) -> Option<i16> {
    u16_of(p).map(|v| v as i16)
}

fn i32_of(p: &Prl) -> Option<i32> {
    match p.operand {
        [b0, b1, b2, b3, ..] => Some(i32::from_le_bytes([*b0, *b1, *b2, *b3])),
        _ => None,
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn relative_depth_is_signed_ordered_and_bounded() {
        let depth = |ops: &[(u16, i32)]| {
            let bytes: Vec<u8> = ops.iter().flat_map(|(op, value)| op.to_le_bytes().into_iter().chain(value.to_le_bytes())).collect();
            decode(&bytes).depth
        };
        assert_eq!(depth(&[(P_ITAP, 1), (P_DTAP, 1)]), 2);
        assert_eq!(depth(&[(P_ITAP, 3), (P_DTAP, -1)]), 2);
        assert_eq!(depth(&[(P_DTAP, 4), (P_ITAP, 1), (P_DTAP, 1)]), 2);
        assert_eq!(depth(&[(P_DTAP, 1), (P_DTAP, 1)]), 2);
        assert_eq!(depth(&[(P_ITAP, 1), (P_DTAP, i32::MIN)]), 0);
        assert_eq!(depth(&[(P_ITAP, i32::MAX), (P_DTAP, i32::MAX), (P_DTAP, 2)]), u32::MAX);
    }

    /// Text beyond the depth cap stays in the deepest cell; returning to an outer
    /// sibling cell or body paragraph must neither consume it early nor lose it.
    #[test]
    fn deep_tables_preserve_text_at_the_cap_and_after_unwinding() {
        let para = |text: &str, row: Option<RowInfo>| {
            let mut p = Paragraph::new();
            p.text = text.into();
            ParaOut { para: p, end_cp: 0, row: row.map(Box::new) }
        };
        let mut input = Vec::new();
        // Start directly at a depth beyond the cap, without intermediate marks.
        for _ in 0..1024 {
            input.push(para("deep", Some(RowInfo { in_table: true, depth: MAX_TABLE_DEPTH + 1, ..Default::default() })));
        }
        for depth in (2..=MAX_TABLE_DEPTH).rev() {
            input.push(para("", Some(RowInfo { in_table: true, depth, inner_ttp: true, ..Default::default() })));
        }
        input.push(para("", Some(RowInfo { in_table: true, depth: 1, cell_mark: true, ..Default::default() })));
        input.push(para("sibling", Some(RowInfo { in_table: true, depth: 1, cell_mark: true, ..Default::default() })));
        input.push(para(
            "",
            Some(RowInfo { in_table: true, depth: 1, cell_mark: true, ttp: true, edges: vec![0, 1440, 2880], ..Default::default() }),
        ));
        input.push(para("tail", None));
        let blocks = assemble(input);
        assert_eq!(blocks.len(), 2);
        assert_eq!(blocks[1].as_para().expect("tail").text, "tail");
        let outer = blocks[0].as_table().expect("outer table");
        assert_eq!(outer.rows.len(), 1);
        assert_eq!(outer.rows[0].cells.len(), 2);
        assert_eq!(outer.rows[0].cells[1].blocks[0].as_para().expect("sibling").text, "sibling");
        let mut cell = &outer.rows[0].cells[0];
        for _ in 2..=MAX_TABLE_DEPTH {
            let table = cell.blocks[0].as_table().expect("nested table");
            assert_eq!(table.rows.len(), 1);
            assert_eq!(table.rows[0].cells.len(), 1);
            cell = &table.rows[0].cells[0];
        }
        assert_eq!(cell.blocks.len(), 1024);
        assert!(cell.blocks.iter().all(|b| b.as_para().is_some_and(|p| p.text == "deep")));
    }
}
