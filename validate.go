package xlfill

import (
	"fmt"
	"strings"

	"github.com/expr-lang/expr"
)

// Severity indicates the severity of a validation issue.
type Severity int

const (
	SeverityError   Severity = iota // Template will fail at runtime
	SeverityWarning                 // Template may produce unexpected results
)

// ValidationIssue represents a single problem found during template validation.
type ValidationIssue struct {
	Severity Severity
	CellRef  CellRef
	Message  string
}

// String formats the issue as "[ERROR] Sheet1!A2: message" or "[WARN] ...".
func (v ValidationIssue) String() string {
	sev := "ERROR"
	if v.Severity == SeverityWarning {
		sev = "WARN"
	}
	return fmt.Sprintf("[%s] %s: %s", sev, v.CellRef, v.Message)
}

// Validate checks a template for structural and expression errors without
// requiring data. It returns a list of issues found. A non-nil error indicates
// the template could not be opened or parsed at all.
func Validate(templatePath string, opts ...Option) ([]ValidationIssue, error) {
	allOpts := append([]Option{WithTemplate(templatePath)}, opts...)
	filler := NewFiller(allOpts...)
	return filler.Validate()
}

// Validate opens the template and performs static validation checks.
// Structural errors (missing jx:area, invalid cell refs) cause a non-nil error return.
// Expression syntax errors and bounds violations are returned as issues.
func (f *Filler) Validate() ([]ValidationIssue, error) {
	tx, err := f.openTemplate()
	if err != nil {
		return nil, err
	}
	defer tx.Close()

	areas, err := f.BuildAreas(tx)
	if err != nil {
		return nil, fmt.Errorf("build areas: %w", err)
	}

	var issues []ValidationIssue
	issues = append(issues, f.validateLastCellBounds(areas)...)
	issues = append(issues, f.validateExpressions(tx, areas)...)
	issues = append(issues, f.validateCommandAttributes(areas)...)
	return issues, nil
}

// validateLastCellBounds checks that every command's area fits within its parent area.
func (f *Filler) validateLastCellBounds(areas []*Area) []ValidationIssue {
	var issues []ValidationIssue
	for _, area := range areas {
		for _, b := range area.Bindings {
			cmdEndRow := b.StartRef.Row + b.Size.Height - 1
			cmdEndCol := b.StartRef.Col + b.Size.Width - 1
			areaEndRow := area.StartCell.Row + area.AreaSize.Height - 1
			areaEndCol := area.StartCell.Col + area.AreaSize.Width - 1

			if cmdEndRow > areaEndRow || cmdEndCol > areaEndCol {
				issues = append(issues, ValidationIssue{
					Severity: SeverityError,
					CellRef:  b.StartRef,
					Message: fmt.Sprintf("command %q lastCell extends beyond parent area (command ends at row %d col %d, area ends at row %d col %d)",
						b.Command.Name(), cmdEndRow+1, cmdEndCol+1, areaEndRow+1, areaEndCol+1),
				})
			}

			// Recurse into child command areas
			if childArea := getCommandArea(b.Command); childArea != nil {
				issues = append(issues, f.validateLastCellBounds([]*Area{childArea})...)
			}
		}
	}
	return issues
}

// validateExpressions checks expression syntax in all cells within areas.
func (f *Filler) validateExpressions(tx Transformer, areas []*Area) []ValidationIssue {
	var issues []ValidationIssue
	notationBegin := f.opts.notationBegin
	notationEnd := f.opts.notationEnd

	for _, area := range areas {
		for row := 0; row < area.AreaSize.Height; row++ {
			for col := 0; col < area.AreaSize.Width; col++ {
				ref := NewCellRef(area.StartCell.Sheet, area.StartCell.Row+row, area.StartCell.Col+col)
				cd := tx.GetCellData(ref)
				if cd == nil {
					continue
				}

				// Check cell value expressions
				if strVal, ok := cd.Value.(string); ok && strings.Contains(strVal, notationBegin) {
					issues = append(issues, checkExpressionSyntax(ref, strVal, notationBegin, notationEnd)...)
				}

				// Check parameterized formula expressions
				if cd.Formula != "" && strings.Contains(cd.Formula, notationBegin) {
					issues = append(issues, checkExpressionSyntax(ref, cd.Formula, notationBegin, notationEnd)...)
				}
			}
		}
	}
	return issues
}

// checkExpressionSyntax extracts ${...} expressions from a string and compiles them for syntax checking.
func checkExpressionSyntax(ref CellRef, value, notationBegin, notationEnd string) []ValidationIssue {
	var issues []ValidationIssue
	segments := ParseExpressions(value, notationBegin, notationEnd)
	for _, seg := range segments {
		if !seg.IsExpression {
			continue
		}
		_, err := expr.Compile(seg.Text, expr.AllowUndefinedVariables())
		if err != nil {
			issues = append(issues, ValidationIssue{
				Severity: SeverityError,
				CellRef:  ref,
				Message:  fmt.Sprintf("invalid expression syntax %q: %v", seg.Text, err),
			})
		}
	}
	return issues
}

// validateCommandAttributes checks that command attribute expressions have valid syntax.
func (f *Filler) validateCommandAttributes(areas []*Area) []ValidationIssue {
	var issues []ValidationIssue
	for _, area := range areas {
		for _, b := range area.Bindings {
			switch cmd := b.Command.(type) {
			case *EachCommand:
				if issue := compileCheck(b.StartRef, "each", "items", cmd.Items); issue != nil {
					issues = append(issues, *issue)
				}
				if cmd.Select != "" {
					if issue := compileCheck(b.StartRef, "each", "select", cmd.Select); issue != nil {
						issues = append(issues, *issue)
					}
				}
			case *IfCommand:
				if issue := compileCheck(b.StartRef, "if", "condition", cmd.Condition); issue != nil {
					issues = append(issues, *issue)
				}
			case *GridCommand:
				if issue := compileCheck(b.StartRef, "grid", "headers", cmd.Headers); issue != nil {
					issues = append(issues, *issue)
				}
				if issue := compileCheck(b.StartRef, "grid", "data", cmd.Data); issue != nil {
					issues = append(issues, *issue)
				}
			}

			// Recurse into child areas
			if childArea := getCommandArea(b.Command); childArea != nil {
				issues = append(issues, f.validateCommandAttributes([]*Area{childArea})...)
			}
		}
	}
	return issues
}

// compileCheck compiles an expression for syntax checking and returns an issue if it fails.
func compileCheck(ref CellRef, cmdName, attrName, expression string) *ValidationIssue {
	if expression == "" {
		return nil
	}
	_, err := expr.Compile(expression, expr.AllowUndefinedVariables())
	if err != nil {
		return &ValidationIssue{
			Severity: SeverityError,
			CellRef:  ref,
			Message:  fmt.Sprintf("%s command has invalid %s expression %q: %v", cmdName, attrName, expression, err),
		}
	}
	return nil
}
