use super::common::create_cell_ref_ast;
use crate::engine::topology::{CellRole, analyze};
use crate::engine::DependencyGraph;
use formualizer_common::LiteralValue;

/// Helper: set a plain value and a formula that references cells.
/// `refs` is a list of (sheet, row, col) that the formula depends on.
fn set_formula_referencing(
    graph: &mut DependencyGraph,
    sheet: &str,
    row: u32,
    col: u32,
    refs: &[(Option<&str>, u32, u32)],
) {
    assert!(!refs.is_empty());
    if refs.len() == 1 {
        let (ref_sheet, r, c) = refs[0];
        let ast = create_cell_ref_ast(ref_sheet, r, c);
        graph.set_cell_formula(sheet, row, col, ast).unwrap();
    } else {
        // Build a binary op chain: ref1 + ref2 + ...
        let mut ast = create_cell_ref_ast(refs[0].0, refs[0].1, refs[0].2);
        for &(ref_sheet, r, c) in &refs[1..] {
            let right = create_cell_ref_ast(ref_sheet, r, c);
            ast = super::common::create_binary_op_ast(ast, right, "+");
        }
        graph.set_cell_formula(sheet, row, col, ast).unwrap();
    }
}

/// Diamond graph: A1→B1, A1→C1, B1→D1, C1→D1
/// A1 = value, B1 = =A1, C1 = =A1, D1 = =B1+C1
///
/// Expected:
/// - A1: key_driver, downstream_reach >= 3
/// - D1: summary_output, upstream_depth = 2
/// - B1, C1: passthrough
#[test]
fn test_diamond_graph() {
    let mut graph = DependencyGraph::new();

    // A1 = value
    graph
        .set_cell_value("Sheet1", 1, 1, LiteralValue::Int(100))
        .unwrap();

    // B1 = =A1
    set_formula_referencing(&mut graph, "Sheet1", 1, 2, &[(Some("Sheet1"), 1, 1)]);

    // C1 = =A1
    set_formula_referencing(&mut graph, "Sheet1", 1, 3, &[(Some("Sheet1"), 1, 1)]);

    // D1 = =B1 + C1
    set_formula_referencing(
        &mut graph,
        "Sheet1",
        1,
        4,
        &[(Some("Sheet1"), 1, 2), (Some("Sheet1"), 1, 3)],
    );

    let result = analyze(&graph, 10);

    // Find cells by address
    let a1 = result
        .cells
        .iter()
        .find(|c| c.address == "A1")
        .expect("A1 should be in results");
    let d1 = result
        .cells
        .iter()
        .find(|c| c.address == "D1")
        .expect("D1 should be in results");

    assert_eq!(a1.classification, CellRole::KeyDriver);
    assert!(
        a1.downstream_reach >= 3,
        "A1 downstream_reach should be >= 3, got {}",
        a1.downstream_reach
    );
    assert_eq!(a1.fan_out, 2); // B1 and C1 depend on A1

    assert_eq!(d1.classification, CellRole::SummaryOutput);
    assert_eq!(d1.upstream_depth, 2);
    assert_eq!(d1.fan_out, 0);

    // B1 and C1 should be passthrough (not in cells since cells only has drivers + outputs)
    // Verify via sheet topology
    let sheet = &result.sheets[0];
    assert_eq!(sheet.name, "Sheet1");
    assert_eq!(sheet.vertex_count, 4);
    assert_eq!(sheet.key_driver_count, 1);
    assert_eq!(sheet.summary_output_count, 1);
}

/// Island detection: standalone cell with no edges.
#[test]
fn test_island_detection() {
    let mut graph = DependencyGraph::new();

    graph
        .set_cell_value("Sheet1", 1, 1, LiteralValue::Int(42))
        .unwrap();

    let result = analyze(&graph, 10);

    // A1 should be an island — no edges, so not in cells list (cells only has drivers + outputs)
    assert!(result.cells.is_empty());
    assert_eq!(result.model.total_islands, 1);
    assert_eq!(result.sheets[0].island_count, 1);
}

/// Cross-sheet: Sheet1!A1 → Sheet2!A1 (Sheet2!A1 formula references Sheet1!A1)
#[test]
fn test_cross_sheet() {
    let mut graph = DependencyGraph::new();

    // Sheet1!A1 = value
    graph
        .set_cell_value("Sheet1", 1, 1, LiteralValue::Int(100))
        .unwrap();

    // Sheet2!A1 = =Sheet1!A1
    set_formula_referencing(&mut graph, "Sheet2", 1, 1, &[(Some("Sheet1"), 1, 1)]);

    let result = analyze(&graph, 10);

    let driver = result
        .cells
        .iter()
        .find(|c| c.classification == CellRole::KeyDriver)
        .expect("should have a key driver");
    assert_eq!(driver.sheet, "Sheet1");
    assert!(driver.is_cross_sheet);

    let output = result
        .cells
        .iter()
        .find(|c| c.classification == CellRole::SummaryOutput)
        .expect("should have a summary output");
    assert_eq!(output.sheet, "Sheet2");
    assert!(output.is_cross_sheet);

    // Sheet edges
    let sheet1 = result.sheets.iter().find(|s| s.name == "Sheet1").unwrap();
    assert!(sheet1.feeds_sheets.contains(&"Sheet2".to_string()));

    let sheet2 = result.sheets.iter().find(|s| s.name == "Sheet2").unwrap();
    assert!(sheet2.fed_by_sheets.contains(&"Sheet1".to_string()));

    // Model sheet_edges
    assert!(result
        .model
        .sheet_edges
        .contains(&("Sheet1".to_string(), "Sheet2".to_string())));
}

/// Linear chain: A1 → B1 → C1 → D1
/// A1 = value, B1 = =A1, C1 = =B1, D1 = =C1
#[test]
fn test_linear_chain() {
    let mut graph = DependencyGraph::new();

    graph
        .set_cell_value("Sheet1", 1, 1, LiteralValue::Int(1))
        .unwrap();

    set_formula_referencing(&mut graph, "Sheet1", 1, 2, &[(Some("Sheet1"), 1, 1)]);
    set_formula_referencing(&mut graph, "Sheet1", 1, 3, &[(Some("Sheet1"), 1, 2)]);
    set_formula_referencing(&mut graph, "Sheet1", 1, 4, &[(Some("Sheet1"), 1, 3)]);

    let result = analyze(&graph, 10);

    let a1 = result
        .cells
        .iter()
        .find(|c| c.address == "A1")
        .expect("A1 should be a key driver");
    assert_eq!(a1.classification, CellRole::KeyDriver);
    assert_eq!(a1.downstream_reach, 3);

    let d1 = result
        .cells
        .iter()
        .find(|c| c.address == "D1")
        .expect("D1 should be a summary output");
    assert_eq!(d1.classification, CellRole::SummaryOutput);
    assert_eq!(d1.upstream_depth, 3);
}

/// Fan-out hub: A1 → B1, C1, D1, E1
#[test]
fn test_fan_out_hub() {
    let mut graph = DependencyGraph::new();

    graph
        .set_cell_value("Sheet1", 1, 1, LiteralValue::Int(10))
        .unwrap();

    // B1..E1 all depend on A1
    for col in 2..=5 {
        set_formula_referencing(&mut graph, "Sheet1", 1, col, &[(Some("Sheet1"), 1, 1)]);
    }

    let result = analyze(&graph, 10);

    let a1 = result
        .cells
        .iter()
        .find(|c| c.address == "A1")
        .expect("A1 should be a key driver");
    assert_eq!(a1.classification, CellRole::KeyDriver);
    assert_eq!(a1.fan_out, 4);
    assert_eq!(a1.downstream_reach, 4);

    // All of B1..E1 should be summary outputs
    let outputs: Vec<_> = result
        .cells
        .iter()
        .filter(|c| c.classification == CellRole::SummaryOutput)
        .collect();
    assert_eq!(outputs.len(), 4);
}

/// Empty graph: no formulas, all vertices are islands
#[test]
fn test_empty_graph() {
    let mut graph = DependencyGraph::new();

    // Add some values but no formulas
    graph
        .set_cell_value("Sheet1", 1, 1, LiteralValue::Int(1))
        .unwrap();
    graph
        .set_cell_value("Sheet1", 2, 1, LiteralValue::Int(2))
        .unwrap();
    graph
        .set_cell_value("Sheet1", 3, 1, LiteralValue::Int(3))
        .unwrap();

    let result = analyze(&graph, 10);

    assert_eq!(result.model.total_vertices, 3);
    assert_eq!(result.model.total_formulas, 0);
    assert_eq!(result.model.total_islands, 3);
    assert!(result.cells.is_empty());
}

/// Truly empty graph: no vertices at all
#[test]
fn test_no_vertices() {
    let graph = DependencyGraph::new();
    let result = analyze(&graph, 10);

    assert_eq!(result.model.total_vertices, 0);
    assert!(result.cells.is_empty());
    assert!(result.sheets.is_empty());
}

/// Test top-N limiting
#[test]
fn test_top_n_limiting() {
    let mut graph = DependencyGraph::new();

    // Create 5 independent value→formula pairs
    for i in 1..=5 {
        graph
            .set_cell_value("Sheet1", i, 1, LiteralValue::Int(i as i64))
            .unwrap();
        set_formula_referencing(
            &mut graph,
            "Sheet1",
            i,
            2,
            &[(Some("Sheet1"), i, 1)],
        );
    }

    // Request top_n = 2
    let result = analyze(&graph, 2);

    assert!(result.sheets[0].top_drivers.len() <= 2);
    assert!(result.sheets[0].top_outputs.len() <= 2);
    assert!(result.model.top_drivers.len() <= 2);
    assert!(result.model.top_outputs.len() <= 2);
}

/// Test sheet topology counts
#[test]
fn test_sheet_topology_counts() {
    let mut graph = DependencyGraph::new();

    // Sheet1: 2 values, 2 formulas, 1 island
    graph
        .set_cell_value("Sheet1", 1, 1, LiteralValue::Int(1))
        .unwrap();
    graph
        .set_cell_value("Sheet1", 2, 1, LiteralValue::Int(2))
        .unwrap();
    graph
        .set_cell_value("Sheet1", 3, 1, LiteralValue::Int(999))
        .unwrap(); // island

    // C1 = =A1 + B1
    set_formula_referencing(
        &mut graph,
        "Sheet1",
        1,
        3,
        &[(Some("Sheet1"), 1, 1), (Some("Sheet1"), 2, 1)],
    );

    let result = analyze(&graph, 10);

    let sheet = &result.sheets[0];
    assert_eq!(sheet.vertex_count, 4);
    assert_eq!(sheet.formula_count, 1);
    assert_eq!(sheet.key_driver_count, 2); // A1, B1
    assert_eq!(sheet.summary_output_count, 1); // C1
    assert_eq!(sheet.island_count, 1); // A3
}
