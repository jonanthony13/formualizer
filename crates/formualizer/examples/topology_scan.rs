//! Scan xlsx files and print topology analysis for each.
//!
//! Usage:
//!   cargo run --example topology_scan --features calamine -- <directory>

use formualizer::workbook::{LoadStrategy, Workbook, WorkbookConfig};
use std::path::Path;
use std::time::Instant;

fn main() {
    let dir = std::env::args()
        .nth(1)
        .unwrap_or_else(|| ".".to_string());
    let dir = Path::new(&dir);

    let mut entries: Vec<_> = std::fs::read_dir(dir)
        .expect("cannot read directory")
        .filter_map(|e| e.ok())
        .filter(|e| {
            let name = e.file_name();
            let name = name.to_string_lossy();
            (name.ends_with(".xlsx") || name.ends_with(".xlsm") || name.ends_with(".xls"))
                && !name.starts_with('~')
        })
        .collect();
    entries.sort_by_key(|e| e.file_name());

    if entries.is_empty() {
        eprintln!("No xlsx/xlsm/xls files found in {}", dir.display());
        std::process::exit(1);
    }

    for entry in &entries {
        let path = entry.path();
        let name = path.file_name().unwrap().to_string_lossy();
        let sep = "=".repeat(70);
        println!("\n{sep}\n{name}\n{sep}");

        let t0 = Instant::now();
        let mut wb = match load_workbook(&path) {
            Ok(wb) => wb,
            Err(e) => {
                println!("  SKIP (load error): {e}");
                continue;
            }
        };
        let load_ms = t0.elapsed().as_millis();

        // Build the dependency graph. Interactive mode defers graph construction
        // until explicitly triggered. We call build_graph_all() to construct edges
        // without running evaluation (which can panic on some models).
        let t_build = Instant::now();
        if let Err(e) = wb.engine_mut().build_graph_all() {
            println!("  WARN (graph build error): {e}");
        }
        let build_ms = t_build.elapsed().as_millis();

        let t1 = Instant::now();
        let topo = wb.engine().analyze_topology(10);
        let topo_ms = t1.elapsed().as_millis();

        println!("  Loaded in {load_ms}ms, graph in {build_ms}ms, topology in {topo_ms}ms");
        println!(
            "  Vertices: {}  Formulas: {}  Islands: {}",
            topo.model.total_vertices, topo.model.total_formulas, topo.model.total_islands
        );

        if !topo.model.sheet_edges.is_empty() {
            println!("  Sheet edges:");
            for (src, tgt) in &topo.model.sheet_edges {
                println!("    {src} -> {tgt}");
            }
        }

        println!("\n  Top drivers (model-wide):");
        if topo.model.top_drivers.is_empty() {
            println!("    (none)");
        }
        for d in &topo.model.top_drivers {
            println!(
                "    {sheet}!{addr}  reach={reach}  fan_out={fo}{cross}",
                sheet = d.sheet,
                addr = d.address,
                reach = d.downstream_reach,
                fo = d.fan_out,
                cross = if d.is_cross_sheet { "  [cross-sheet]" } else { "" },
            );
        }

        println!("\n  Top outputs (model-wide):");
        if topo.model.top_outputs.is_empty() {
            println!("    (none)");
        }
        for o in &topo.model.top_outputs {
            println!(
                "    {sheet}!{addr}  depth={depth}  fan_in={fi}{cross}",
                sheet = o.sheet,
                addr = o.address,
                depth = o.upstream_depth,
                fi = o.fan_in,
                cross = if o.is_cross_sheet { "  [cross-sheet]" } else { "" },
            );
        }

        println!("\n  Per-sheet summary:");
        for s in &topo.sheets {
            println!(
                "    {name}: {v} vertices, {f} formulas, {d} drivers, {o} outputs, {i} islands",
                name = s.name,
                v = s.vertex_count,
                f = s.formula_count,
                d = s.key_driver_count,
                o = s.summary_output_count,
                i = s.island_count,
            );
            if !s.feeds_sheets.is_empty() {
                println!("      feeds: {}", s.feeds_sheets.join(", "));
            }
            if !s.fed_by_sheets.is_empty() {
                println!("      fed by: {}", s.fed_by_sheets.join(", "));
            }
        }
    }
    println!();
}

fn load_workbook(path: &Path) -> Result<Workbook, String> {
    use formualizer::workbook::backends::CalamineAdapter;
    use formualizer::workbook::traits::SpreadsheetReader;

    // Use interactive mode: formulas are staged as text during load, then
    // built into the graph all-at-once (single CSR build across all sheets).
    let config = WorkbookConfig::interactive();
    let adapter = <CalamineAdapter as SpreadsheetReader>::open_path(path)
        .map_err(|e| format!("{e}"))?;
    let wb = Workbook::from_reader(adapter, LoadStrategy::EagerAll, config)
        .map_err(|e| format!("{e}"))?;
    Ok(wb)
}
