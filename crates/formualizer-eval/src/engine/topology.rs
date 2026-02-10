//! Graph-based heuristics for cell classification.
//!
//! Performs a single-pass O(V+E) analysis of the dependency graph, producing
//! per-cell and per-sheet topology metrics: downstream reach, upstream depth,
//! fan-in, fan-out, and a role classification (key driver, summary output,
//! passthrough, island).

use super::graph::DependencyGraph;
use super::vertex::VertexKind;
use formualizer_common::Coord as AbsCoord;
use rustc_hash::{FxHashMap, FxHashSet};
use std::collections::VecDeque;

// ---------------------------------------------------------------------------
// Public data structures
// ---------------------------------------------------------------------------

/// Classification of a cell's structural role in the calculation graph.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum CellRole {
    /// Plain value cell with downstream dependents — a key model input.
    KeyDriver,
    /// Formula cell with no dependents — a terminal output.
    SummaryOutput,
    /// Formula cell that both consumes and produces — an intermediate calc.
    Passthrough,
    /// Cell with no edges in either direction — disconnected.
    Island,
}

impl CellRole {
    pub fn as_str(&self) -> &'static str {
        match self {
            CellRole::KeyDriver => "key_driver",
            CellRole::SummaryOutput => "summary_output",
            CellRole::Passthrough => "passthrough",
            CellRole::Island => "island",
        }
    }
}

/// Per-cell topology metrics.
#[derive(Debug, Clone)]
pub struct CellTopology {
    pub sheet: String,
    pub row: u32,
    pub col: u32,
    pub address: String,
    pub classification: CellRole,
    /// Primary score: downstream_reach for drivers, upstream_depth for outputs.
    pub score: f64,
    /// Number of cells that directly depend on this cell.
    pub fan_out: u32,
    /// Number of cells this cell directly depends on.
    pub fan_in: u32,
    /// Transitive count of cells affected downstream (capped at u32::MAX).
    pub downstream_reach: u32,
    /// Longest dependency chain depth upstream.
    pub upstream_depth: u32,
    /// Whether this cell has a cross-sheet dependency (in or out).
    pub is_cross_sheet: bool,
    /// Whether this cell contains a formula.
    pub has_formula: bool,
}

/// Per-sheet topology summary.
#[derive(Debug, Clone)]
pub struct SheetTopology {
    pub name: String,
    pub vertex_count: u32,
    pub formula_count: u32,
    pub key_driver_count: u32,
    pub summary_output_count: u32,
    pub island_count: u32,
    /// Top-N key drivers by downstream_reach.
    pub top_drivers: Vec<CellTopology>,
    /// Top-N summary outputs by upstream_depth.
    pub top_outputs: Vec<CellTopology>,
    /// Sheet names that this sheet feeds into.
    pub feeds_sheets: Vec<String>,
    /// Sheet names that feed into this sheet.
    pub fed_by_sheets: Vec<String>,
}

/// Model-level topology summary.
#[derive(Debug, Clone)]
pub struct ModelTopology {
    pub total_vertices: u32,
    pub total_formulas: u32,
    pub total_islands: u32,
    pub top_drivers: Vec<CellTopology>,
    pub top_outputs: Vec<CellTopology>,
    /// Directed (source_sheet, target_sheet) pairs.
    pub sheet_edges: Vec<(String, String)>,
}

/// Complete topology analysis result.
#[derive(Debug, Clone)]
pub struct TopologyAnalysis {
    /// Filtered to key_drivers + summary_outputs only.
    pub cells: Vec<CellTopology>,
    pub sheets: Vec<SheetTopology>,
    pub model: ModelTopology,
}

// ---------------------------------------------------------------------------
// Core analysis
// ---------------------------------------------------------------------------

/// Analyze the dependency graph topology. `top_n` controls how many top
/// drivers/outputs are returned per sheet and globally.
pub fn analyze(graph: &DependencyGraph, top_n: usize) -> TopologyAnalysis {
    // Collect all active vertices
    let vids: Vec<_> = graph.iter_vertex_ids().collect();
    let n = vids.len();
    if n == 0 {
        return empty_result();
    }

    // Map VertexId → dense index for O(1) array access
    let mut vid_to_idx: FxHashMap<super::vertex::VertexId, usize> =
        FxHashMap::with_capacity_and_hasher(n, Default::default());
    for (i, &vid) in vids.iter().enumerate() {
        vid_to_idx.insert(vid, i);
    }

    // Read per-vertex metadata
    let store = graph.store_ref();
    let sheet_reg = graph.sheet_reg();
    let edges = graph.edges_ref();

    let mut kinds = Vec::with_capacity(n);
    let mut sheet_ids = Vec::with_capacity(n);
    let mut coords: Vec<AbsCoord> = Vec::with_capacity(n);
    let mut has_formula = Vec::with_capacity(n);

    for &vid in &vids {
        let kind = store.kind(vid);
        kinds.push(kind);
        sheet_ids.push(store.sheet_id(vid));
        coords.push(store.coord(vid));
        has_formula.push(matches!(
            kind,
            VertexKind::FormulaScalar | VertexKind::FormulaArray
        ));
    }

    // Build local adjacency: deps[i] = indices this vertex depends on,
    //                         rdeps[i] = indices that depend on this vertex.
    // Also compute fan_in (number of dependencies) and fan_out (number of dependents).
    let mut fan_in = vec![0u32; n]; // count of dependencies (out_edges)
    let mut fan_out = vec![0u32; n]; // count of dependents (in_edges)

    // deps_of[i] = list of dense indices that vertex i depends on
    let mut deps_of: Vec<Vec<usize>> = vec![Vec::new(); n];
    // rdeps_of[i] = list of dense indices that depend on vertex i
    let mut rdeps_of: Vec<Vec<usize>> = vec![Vec::new(); n];

    for (i, &vid) in vids.iter().enumerate() {
        // out_edges(vid) returns dependencies of vid
        let out = edges.out_edges(vid);
        fan_in[i] = out.len() as u32;
        for dep_vid in &out {
            if let Some(&j) = vid_to_idx.get(dep_vid) {
                deps_of[i].push(j);
                rdeps_of[j].push(i);
            }
        }
    }
    // fan_out = number of reverse dependents
    for i in 0..n {
        fan_out[i] = rdeps_of[i].len() as u32;
    }

    // ------------------------------------------------------------------
    // Kahn's algorithm: topological sort
    // ------------------------------------------------------------------
    // "in-degree" for Kahn's = number of dependencies (fan_in).
    // Sources = vertices with no dependencies.
    let mut in_deg: Vec<u32> = fan_in.clone();
    let mut queue: VecDeque<usize> = VecDeque::new();
    for i in 0..n {
        if in_deg[i] == 0 {
            queue.push_back(i);
        }
    }
    let mut topo_order: Vec<usize> = Vec::with_capacity(n);
    while let Some(i) = queue.pop_front() {
        topo_order.push(i);
        for &j in &rdeps_of[i] {
            in_deg[j] -= 1;
            if in_deg[j] == 0 {
                queue.push_back(j);
            }
        }
    }
    // Vertices not in topo_order are in cycles — they get depth=0, reach=0.
    let mut in_topo = vec![false; n];
    for &i in &topo_order {
        in_topo[i] = true;
    }

    // ------------------------------------------------------------------
    // Pass 1 — forward topo order: upstream_depth
    // upstream_depth[v] = max(upstream_depth[dep] + 1) over all deps
    // ------------------------------------------------------------------
    let mut upstream_depth = vec![0u32; n];
    for &i in &topo_order {
        for &dep in &deps_of[i] {
            let candidate = upstream_depth[dep].saturating_add(1);
            if candidate > upstream_depth[i] {
                upstream_depth[i] = candidate;
            }
        }
    }

    // ------------------------------------------------------------------
    // Pass 2 — reverse topo order: downstream_reach
    // downstream_reach[v] = 1 + sum(downstream_reach[d]) for each dependent d
    // Capped at u32::MAX via saturating arithmetic; tie-break on fan_out.
    // ------------------------------------------------------------------
    let mut downstream_reach = vec![0u32; n];
    for &i in topo_order.iter().rev() {
        for &d in &rdeps_of[i] {
            downstream_reach[i] = downstream_reach[i].saturating_add(1u32.saturating_add(downstream_reach[d]));
        }
    }

    // ------------------------------------------------------------------
    // Cross-sheet detection
    // ------------------------------------------------------------------
    let mut is_cross_sheet = vec![false; n];
    for (i, &vid) in vids.iter().enumerate() {
        let my_sheet = sheet_ids[i];
        let out = edges.out_edges(vid);
        for dep_vid in &out {
            if let Some(&j) = vid_to_idx.get(dep_vid) {
                if sheet_ids[j] != my_sheet {
                    is_cross_sheet[i] = true;
                    is_cross_sheet[j] = true;
                }
            }
        }
    }

    // ------------------------------------------------------------------
    // Classify each vertex
    // ------------------------------------------------------------------
    let mut classifications = Vec::with_capacity(n);
    let mut scores = Vec::with_capacity(n);
    for i in 0..n {
        let (role, score) = classify(
            kinds[i],
            has_formula[i],
            fan_in[i],
            fan_out[i],
            downstream_reach[i],
            upstream_depth[i],
            in_topo[i],
        );
        classifications.push(role);
        scores.push(score);
    }

    // ------------------------------------------------------------------
    // Build CellTopology entries and collect per-sheet data
    // ------------------------------------------------------------------
    let mut all_cells: Vec<CellTopology> = Vec::with_capacity(n);
    // sheet_name -> indices into all_cells
    let mut sheet_cell_indices: FxHashMap<String, Vec<usize>> = FxHashMap::default();

    for i in 0..n {
        let sheet_name = sheet_reg.name(sheet_ids[i]).to_string();
        // Coords are 0-based internally; convert to 1-based for display
        let row_1 = coords[i].row() + 1;
        let col_1 = coords[i].col() + 1;
        let address = format_a1(col_1, row_1);

        let ct = CellTopology {
            sheet: sheet_name.clone(),
            row: row_1,
            col: col_1,
            address,
            classification: classifications[i],
            score: scores[i],
            fan_out: fan_out[i],
            fan_in: fan_in[i],
            downstream_reach: downstream_reach[i],
            upstream_depth: upstream_depth[i],
            is_cross_sheet: is_cross_sheet[i],
            has_formula: has_formula[i],
        };
        let idx = all_cells.len();
        all_cells.push(ct);
        sheet_cell_indices
            .entry(sheet_name)
            .or_default()
            .push(idx);
    }

    // ------------------------------------------------------------------
    // Build sheet-level summaries
    // ------------------------------------------------------------------
    let mut sheet_topos: Vec<SheetTopology> = Vec::new();
    // Cross-sheet edge set
    let mut sheet_edge_set: FxHashSet<(String, String)> = FxHashSet::default();

    // Collect cross-sheet edges
    for (i, &vid) in vids.iter().enumerate() {
        let my_sheet = sheet_ids[i];
        let out = edges.out_edges(vid);
        for dep_vid in &out {
            if let Some(&j) = vid_to_idx.get(dep_vid) {
                if sheet_ids[j] != my_sheet {
                    // This vertex (i) depends on vertex (j) from another sheet.
                    // Data flows from j's sheet into i's sheet.
                    let from_sheet = sheet_reg.name(sheet_ids[j]).to_string();
                    let to_sheet = sheet_reg.name(my_sheet).to_string();
                    sheet_edge_set.insert((from_sheet, to_sheet));
                }
            }
        }
    }

    for (sheet_name, indices) in &sheet_cell_indices {
        let mut vertex_count = 0u32;
        let mut formula_count = 0u32;
        let mut key_driver_count = 0u32;
        let mut summary_output_count = 0u32;
        let mut island_count = 0u32;
        let mut drivers: Vec<&CellTopology> = Vec::new();
        let mut outputs: Vec<&CellTopology> = Vec::new();

        for &idx in indices {
            vertex_count += 1;
            let cell = &all_cells[idx];
            if cell.has_formula {
                formula_count += 1;
            }
            match cell.classification {
                CellRole::KeyDriver => {
                    key_driver_count += 1;
                    drivers.push(cell);
                }
                CellRole::SummaryOutput => {
                    summary_output_count += 1;
                    outputs.push(cell);
                }
                CellRole::Island => {
                    island_count += 1;
                }
                CellRole::Passthrough => {}
            }
        }

        drivers.sort_by(|a, b| {
            b.downstream_reach
                .cmp(&a.downstream_reach)
                .then_with(|| b.fan_out.cmp(&a.fan_out))
                .then_with(|| a.row.cmp(&b.row))
                .then_with(|| a.col.cmp(&b.col))
        });
        outputs.sort_by(|a, b| {
            b.upstream_depth
                .cmp(&a.upstream_depth)
                .then_with(|| b.fan_in.cmp(&a.fan_in))
                .then_with(|| a.row.cmp(&b.row))
                .then_with(|| a.col.cmp(&b.col))
        });

        let top_drivers: Vec<CellTopology> =
            drivers.iter().take(top_n).map(|c| (*c).clone()).collect();
        let top_outputs: Vec<CellTopology> =
            outputs.iter().take(top_n).map(|c| (*c).clone()).collect();

        // sheets this one feeds into (sorted for determinism)
        let mut feeds_sheets: Vec<String> = sheet_edge_set
            .iter()
            .filter(|(from, _)| from == sheet_name)
            .map(|(_, to)| to.clone())
            .collect();
        feeds_sheets.sort();
        // sheets that feed into this one (sorted for determinism)
        let mut fed_by_sheets: Vec<String> = sheet_edge_set
            .iter()
            .filter(|(_, to)| to == sheet_name)
            .map(|(from, _)| from.clone())
            .collect();
        fed_by_sheets.sort();

        sheet_topos.push(SheetTopology {
            name: sheet_name.clone(),
            vertex_count,
            formula_count,
            key_driver_count,
            summary_output_count,
            island_count,
            top_drivers,
            top_outputs,
            feeds_sheets,
            fed_by_sheets,
        });
    }
    sheet_topos.sort_by(|a, b| a.name.cmp(&b.name));

    // ------------------------------------------------------------------
    // Build model-level summary
    // ------------------------------------------------------------------
    let total_vertices = n as u32;
    let total_formulas = has_formula.iter().filter(|&&f| f).count() as u32;
    let total_islands = classifications
        .iter()
        .filter(|&&c| c == CellRole::Island)
        .count() as u32;

    let mut global_drivers: Vec<&CellTopology> = all_cells
        .iter()
        .filter(|c| c.classification == CellRole::KeyDriver)
        .collect();
    global_drivers.sort_by(|a, b| {
        b.downstream_reach
            .cmp(&a.downstream_reach)
            .then_with(|| b.fan_out.cmp(&a.fan_out))
            .then_with(|| a.sheet.cmp(&b.sheet))
            .then_with(|| a.row.cmp(&b.row))
            .then_with(|| a.col.cmp(&b.col))
    });

    let mut global_outputs: Vec<&CellTopology> = all_cells
        .iter()
        .filter(|c| c.classification == CellRole::SummaryOutput)
        .collect();
    global_outputs.sort_by(|a, b| {
        b.upstream_depth
            .cmp(&a.upstream_depth)
            .then_with(|| b.fan_in.cmp(&a.fan_in))
            .then_with(|| a.sheet.cmp(&b.sheet))
            .then_with(|| a.row.cmp(&b.row))
            .then_with(|| a.col.cmp(&b.col))
    });

    let model_top_drivers: Vec<CellTopology> = global_drivers
        .iter()
        .take(top_n)
        .map(|c| (*c).clone())
        .collect();
    let model_top_outputs: Vec<CellTopology> = global_outputs
        .iter()
        .take(top_n)
        .map(|c| (*c).clone())
        .collect();

    let mut sheet_edges: Vec<(String, String)> = sheet_edge_set.into_iter().collect();
    sheet_edges.sort();

    let model = ModelTopology {
        total_vertices,
        total_formulas,
        total_islands,
        top_drivers: model_top_drivers,
        top_outputs: model_top_outputs,
        sheet_edges,
    };

    // Filter cells to key_drivers + summary_outputs only
    let cells: Vec<CellTopology> = all_cells
        .into_iter()
        .filter(|c| {
            matches!(
                c.classification,
                CellRole::KeyDriver | CellRole::SummaryOutput
            )
        })
        .collect();

    TopologyAnalysis {
        cells,
        sheets: sheet_topos,
        model,
    }
}

// ---------------------------------------------------------------------------
// Helpers
// ---------------------------------------------------------------------------

fn classify(
    kind: VertexKind,
    has_formula: bool,
    fan_in: u32,
    fan_out: u32,
    downstream_reach: u32,
    upstream_depth: u32,
    in_topo: bool,
) -> (CellRole, f64) {
    if !in_topo {
        // Vertex in a cycle — treat as island
        return (CellRole::Island, 0.0);
    }
    if fan_in == 0 && fan_out == 0 {
        return (CellRole::Island, 0.0);
    }
    // Key driver: plain value (Cell or Empty kind) with dependents
    if !has_formula && fan_out > 0 {
        return (CellRole::KeyDriver, downstream_reach as f64);
    }
    // Summary output: formula with no dependents
    if has_formula && fan_out == 0 {
        return (CellRole::SummaryOutput, upstream_depth as f64);
    }
    // Passthrough: formula with both inputs and outputs
    if has_formula && fan_in > 0 && fan_out > 0 {
        return (CellRole::Passthrough, downstream_reach as f64);
    }
    // Fallback: non-formula cell with no dependents but has dependencies
    // (shouldn't normally happen, but handle gracefully)
    if fan_out == 0 {
        return (CellRole::Island, 0.0);
    }
    (CellRole::Passthrough, downstream_reach as f64)
}

/// Format a 1-based (col, row) as A1 notation.
fn format_a1(col_1based: u32, row_1based: u32) -> String {
    let col_letters = col_to_letters_0based(col_1based - 1);
    format!("{}{}", col_letters, row_1based)
}

/// Convert 0-based column index to Excel letter notation (0=A, 25=Z, 26=AA).
fn col_to_letters_0based(mut col: u32) -> String {
    let mut buf = Vec::new();
    loop {
        let rem = (col % 26) as u8;
        buf.push(b'A' + rem);
        col /= 26;
        if col == 0 {
            break;
        }
        col -= 1;
    }
    buf.reverse();
    String::from_utf8(buf).expect("only ASCII A-Z")
}

fn empty_result() -> TopologyAnalysis {
    TopologyAnalysis {
        cells: Vec::new(),
        sheets: Vec::new(),
        model: ModelTopology {
            total_vertices: 0,
            total_formulas: 0,
            total_islands: 0,
            top_drivers: Vec::new(),
            top_outputs: Vec::new(),
            sheet_edges: Vec::new(),
        },
    }
}
