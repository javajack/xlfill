package xlfill

import (
	"fmt"
	"regexp"
	"strings"
)

// FormulaProcessor updates formula cell references after template expansion.
type FormulaProcessor interface {
	ProcessAreaFormulas(transformer Transformer, area *Area)
}

// StandardFormulaProcessor implements the standard formula processing algorithm.
// It maps source cell references in formulas to their expanded target positions.
type StandardFormulaProcessor struct{}

// NewFormulaProcessor creates a new StandardFormulaProcessor.
func NewFormulaProcessor() *StandardFormulaProcessor {
	return &StandardFormulaProcessor{}
}

// cellRefRegex matches cell references in formulas (e.g., A1, $A$1, Sheet1!A1, A1:B5).
var cellRefRegex = regexp.MustCompile(`(?:('?[^'!]+?'?)!)?\$?([A-Z]{1,3})\$?(\d+)`)

// ProcessAreaFormulas processes all formula cells in the area, updating references.
func (fp *StandardFormulaProcessor) ProcessAreaFormulas(transformer Transformer, area *Area) {
	formulaCells := transformer.GetFormulaCells()

	for _, cd := range formulaCells {
		if !area.containsRef(cd.Ref) {
			continue
		}

		targetPositions := transformer.GetTargetCellRef(cd.Ref)
		if len(targetPositions) == 0 {
			continue
		}

		for _, targetPos := range targetPositions {
			newFormula := fp.processFormula(cd.Formula, cd, targetPos, transformer, area)
			if newFormula != "" {
				transformer.SetFormula(targetPos, newFormula)
			}
		}
	}
}

// processFormula processes a single formula, replacing source refs with target refs.
func (fp *StandardFormulaProcessor) processFormula(
	formula string,
	formulaCell *CellData,
	targetPos CellRef,
	transformer Transformer,
	area *Area,
) string {
	result := formula

	// Find all cell reference matches in the formula
	matches := cellRefRegex.FindAllStringSubmatchIndex(formula, -1)
	if len(matches) == 0 {
		return formula
	}

	// Process matches in reverse order to preserve indices
	for i := len(matches) - 1; i >= 0; i-- {
		match := matches[i]
		fullMatch := formula[match[0]:match[1]]

		// Parse the referenced cell
		ref, err := parseCellRefFromFormula(fullMatch, area.StartCell.Sheet)
		if err != nil {
			continue
		}

		// Look up where this source cell was mapped to
		targetRefs := transformer.GetTargetCellRef(ref)
		if len(targetRefs) == 0 {
			// External reference — check if it's outside the area
			if !area.containsRef(ref) {
				continue // keep external ref as-is
			}
			// Internal ref with no target — use default value
			defaultVal := formulaCell.DefaultValue
			if defaultVal == "" {
				defaultVal = "0"
			}
			result = result[:match[0]] + defaultVal + result[match[1]:]
			continue
		}

		// Apply formula strategy filtering
		filtered := fp.filterByStrategy(targetRefs, targetPos, formulaCell.FormulaStrategy)
		if len(filtered) == 0 {
			defaultVal := formulaCell.DefaultValue
			if defaultVal == "" {
				defaultVal = "0"
			}
			result = result[:match[0]] + defaultVal + result[match[1]:]
			continue
		}

		// Replace the reference
		replacement := fp.buildReplacement(filtered, ref.Sheet, area.StartCell.Sheet)
		result = result[:match[0]] + replacement + result[match[1]:]
	}

	return result
}

// filterByStrategy filters target refs based on FormulaStrategy.
func (fp *StandardFormulaProcessor) filterByStrategy(
	targets []CellRef, formulaTarget CellRef, strategy FormulaStrategy,
) []CellRef {
	switch strategy {
	case FormulaByColumn:
		var filtered []CellRef
		for _, t := range targets {
			if t.Col == formulaTarget.Col {
				filtered = append(filtered, t)
			}
		}
		return filtered
	case FormulaByRow:
		var filtered []CellRef
		for _, t := range targets {
			if t.Row == formulaTarget.Row {
				filtered = append(filtered, t)
			}
		}
		return filtered
	default:
		return targets
	}
}

// buildReplacement builds the replacement string for a set of target refs.
func (fp *StandardFormulaProcessor) buildReplacement(targets []CellRef, refSheet, areaSheet string) string {
	if len(targets) == 1 {
		return fp.formatRef(targets[0], refSheet, areaSheet)
	}

	// Multiple targets — check if they form a contiguous range
	if rangeStr := fp.tryBuildRange(targets, refSheet, areaSheet); rangeStr != "" {
		return rangeStr
	}

	// Non-contiguous — join with commas (for SUM etc.)
	parts := make([]string, len(targets))
	for i, t := range targets {
		parts[i] = fp.formatRef(t, refSheet, areaSheet)
	}

	// Excel SUM limit: 255 args. If exceeded, use addition chain.
	if len(parts) > 255 {
		return strings.Join(parts, "+")
	}
	return strings.Join(parts, ",")
}

// tryBuildRange checks if targets form a contiguous vertical or horizontal range.
func (fp *StandardFormulaProcessor) tryBuildRange(targets []CellRef, refSheet, areaSheet string) string {
	if len(targets) < 2 {
		return ""
	}

	// Check vertical range (same col, consecutive rows)
	allSameCol := true
	for _, t := range targets {
		if t.Col != targets[0].Col || t.Sheet != targets[0].Sheet {
			allSameCol = false
			break
		}
	}
	if allSameCol {
		minRow, maxRow := targets[0].Row, targets[0].Row
		for _, t := range targets[1:] {
			if t.Row < minRow {
				minRow = t.Row
			}
			if t.Row > maxRow {
				maxRow = t.Row
			}
		}
		if maxRow-minRow+1 == len(targets) {
			first := NewCellRef(targets[0].Sheet, minRow, targets[0].Col)
			last := NewCellRef(targets[0].Sheet, maxRow, targets[0].Col)
			return fp.formatRef(first, refSheet, areaSheet) + ":" + fp.formatRef(last, refSheet, areaSheet)
		}
	}

	// Check horizontal range (same row, consecutive cols)
	allSameRow := true
	for _, t := range targets {
		if t.Row != targets[0].Row || t.Sheet != targets[0].Sheet {
			allSameRow = false
			break
		}
	}
	if allSameRow {
		minCol, maxCol := targets[0].Col, targets[0].Col
		for _, t := range targets[1:] {
			if t.Col < minCol {
				minCol = t.Col
			}
			if t.Col > maxCol {
				maxCol = t.Col
			}
		}
		if maxCol-minCol+1 == len(targets) {
			first := NewCellRef(targets[0].Sheet, targets[0].Row, minCol)
			last := NewCellRef(targets[0].Sheet, targets[0].Row, maxCol)
			return fp.formatRef(first, refSheet, areaSheet) + ":" + fp.formatRef(last, refSheet, areaSheet)
		}
	}

	return ""
}

// formatRef formats a cell reference, adding sheet prefix if needed.
func (fp *StandardFormulaProcessor) formatRef(ref CellRef, origRefSheet, areaSheet string) string {
	cellName := ref.CellName()
	// Add sheet prefix if the reference was cross-sheet or if target is on different sheet
	if origRefSheet != "" && origRefSheet != areaSheet {
		return origRefSheet + "!" + cellName
	}
	if ref.Sheet != "" && ref.Sheet != areaSheet {
		return ref.Sheet + "!" + cellName
	}
	return cellName
}

// parseCellRefFromFormula parses a cell reference from a formula match.
func parseCellRefFromFormula(match string, defaultSheet string) (CellRef, error) {
	// Remove $ signs for parsing
	clean := strings.ReplaceAll(match, "$", "")

	if strings.Contains(clean, "!") {
		return ParseCellRef(clean)
	}

	ref, err := ParseCellRef(clean)
	if err != nil {
		return CellRef{}, err
	}
	ref.Sheet = defaultSheet
	return ref, nil
}

// ProcessFormulasForRange is a convenience method to handle range references in formulas.
// It expands "SUM(C2:C2)" to "SUM(C2:C5)" when C2 was replicated to C2,C3,C4,C5.
func (fp *StandardFormulaProcessor) ProcessFormulasForRange(
	formula string,
	transformer Transformer,
	defaultSheet string,
) string {
	// Find range patterns like A1:B2
	rangeRegex := regexp.MustCompile(`(?:('?[^'!]+?'?)!)?\$?([A-Z]{1,3})\$?(\d+):\$?([A-Z]{1,3})\$?(\d+)`)

	return rangeRegex.ReplaceAllStringFunc(formula, func(match string) string {
		parts := rangeRegex.FindStringSubmatch(match)
		if len(parts) < 6 {
			return match
		}

		sheet := parts[1]
		if sheet == "" {
			sheet = defaultSheet
		}

		startCellStr := fmt.Sprintf("%s%s", parts[2], parts[3])
		endCellStr := fmt.Sprintf("%s%s", parts[4], parts[5])

		startRef, err1 := ParseCellRef(sheet + "!" + startCellStr)
		endRef, err2 := ParseCellRef(sheet + "!" + endCellStr)
		if err1 != nil || err2 != nil {
			return match
		}

		// Check if start cell was expanded
		startTargets := transformer.GetTargetCellRef(startRef)
		endTargets := transformer.GetTargetCellRef(endRef)

		if len(startTargets) == 0 && len(endTargets) == 0 {
			return match // no expansion
		}

		// Find the actual min/max of all targets
		var allTargets []CellRef
		allTargets = append(allTargets, startTargets...)
		allTargets = append(allTargets, endTargets...)

		if len(allTargets) == 0 {
			return match
		}

		minRow, maxRow := allTargets[0].Row, allTargets[0].Row
		minCol, maxCol := allTargets[0].Col, allTargets[0].Col
		for _, t := range allTargets[1:] {
			if t.Row < minRow {
				minRow = t.Row
			}
			if t.Row > maxRow {
				maxRow = t.Row
			}
			if t.Col < minCol {
				minCol = t.Col
			}
			if t.Col > maxCol {
				maxCol = t.Col
			}
		}

		newStart := NewCellRef(sheet, minRow, minCol)
		newEnd := NewCellRef(sheet, maxRow, maxCol)
		return newStart.CellName() + ":" + newEnd.CellName()
	})
}
