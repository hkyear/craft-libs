//! Local diagnostic: dump the table structure doc-binary builds from a .doc file.
use std::collections::VecDeque;
use std::sync::Arc;

fn walk_blocks(blocks: &[Arc<craft_wordmodel::Block>], depth: usize, stats: &mut Vec<(usize, usize, usize, usize, bool, bool)>, out: &mut VecDeque<String>) {
    for b in blocks {
        if let Some(t) = b.as_table() {
            let mut sum_spans = 0u32;
            let mut n_width = 0usize;
            for r in &t.rows {
                for c in &r.cells {
                    sum_spans += c.props.span;
                    if c.props.width.is_some() || c.props.width_pct.is_some() { n_width += 1; }
                }
            }
            stats.push((t.grid.len(), t.rows.len(), t.rows.first().map_or(0, |r| r.cells.len()), sum_spans as usize, n_width > 0, t.props.fixed));
            out.push_back(format!("{}TABLE grid={} rows={} fixed={} first_row_cells={} sum_spans={}", "  ".repeat(depth), t.grid.len(), t.rows.len(), t.props.fixed, t.rows.first().map_or(0, |r| r.cells.len()), sum_spans));
            if t.grid.len() > 0 && t.grid.len() <= 12 {
                out.push_back(format!("{}  grid: {:?}", "  ".repeat(depth), t.grid));
            }
            if t.rows.len() <= 3 {
                for (ri, r) in t.rows.iter().enumerate() {
                    out.push_back(format!("{}  row{} cells={}", "  ".repeat(depth), ri, r.cells.len()));
                }
            }
            for r in &t.rows {
                for c in &r.cells {
                    walk_blocks(&c.blocks, depth + 1, stats, out);
                }
            }
        }
    }
}

fn main() {
    let path = std::env::args().nth(1).expect("usage: probe_doc <file.doc>");
    let bytes = std::fs::read(&path).expect("read");
    let doc = craft_doc_binary::read(&bytes).expect("parse");
    let mut stats = Vec::new();
    let mut out = VecDeque::new();
    walk_blocks(&doc.body, 0, &mut stats, &mut out);
    for l in out { println!("{l}"); }
    println!("=== summary: {} tables ===", stats.len());
    for (i, (g, r, fc, ss, w, fx)) in stats.iter().enumerate() {
        println!("T{i}: grid={g} rows={r} first_cells={fc} sum_spans={ss} has_width={w} fixed={fx}");
    }
}
