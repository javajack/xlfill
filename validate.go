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

// ValidateData checks if the provided data satisfies the template's expression requirements.
// It parses all ${...} expressions, extracts required top-level variable names,
// cross-references them against the data map and command-provided variables (var/varIndex),
// and reports any mismatches.
func (f *Filler) ValidateData(data map[string]any) ([]ValidationIssue, error) {
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

	// Collect variables provided by commands (each var, each varIndex, etc.)
	commandVars := map[string]CellRef{} // variable name → defining command cell
	collectCommandVars(areas, commandVars)

	// Built-in variables always available
	builtins := map[string]bool{
		"_row": true, "_col": true, "hyperlink": true,
		"true": true, "false": true, "nil": true, "len": true,
	}

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
					issues = append(issues, checkDataRequirements(ref, strVal, notationBegin, notationEnd, data, commandVars, builtins)...)
				}

				// Check parameterized formula expressions
				if cd.Formula != "" && strings.Contains(cd.Formula, notationBegin) {
					issues = append(issues, checkDataRequirements(ref, cd.Formula, notationBegin, notationEnd, data, commandVars, builtins)...)
				}
			}
		}
	}

	return issues, nil
}

// collectCommandVars recursively collects variable names provided by commands.
func collectCommandVars(areas []*Area, vars map[string]CellRef) {
	for _, area := range areas {
		for _, b := range area.Bindings {
			switch cmd := b.Command.(type) {
			case *EachCommand:
				vars[cmd.Var] = b.StartRef
				if cmd.VarIndex != "" {
					vars[cmd.VarIndex] = b.StartRef
				}
				if cmd.Area != nil {
					collectCommandVars([]*Area{cmd.Area}, vars)
				}
			case *RepeatCommand:
				if cmd.Var != "" {
					vars[cmd.Var] = b.StartRef
				}
				if cmd.Area != nil {
					collectCommandVars([]*Area{cmd.Area}, vars)
				}
			case *IfCommand:
				if cmd.IfArea != nil {
					collectCommandVars([]*Area{cmd.IfArea}, vars)
				}
				if cmd.ElseArea != nil {
					collectCommandVars([]*Area{cmd.ElseArea}, vars)
				}
			case *GridCommand:
				if cmd.BodyArea != nil {
					collectCommandVars([]*Area{cmd.BodyArea}, vars)
				}
			case *UpdateCellCommand:
				if cmd.Area != nil {
					collectCommandVars([]*Area{cmd.Area}, vars)
				}
			}
		}
	}
}

// checkDataRequirements extracts variable names from expressions and checks they exist in data or command scope.
func checkDataRequirements(ref CellRef, value, notationBegin, notationEnd string, data map[string]any, commandVars map[string]CellRef, builtins map[string]bool) []ValidationIssue {
	var issues []ValidationIssue
	segments := ParseExpressions(value, notationBegin, notationEnd)
	for _, seg := range segments {
		if !seg.IsExpression {
			continue
		}
		// Extract top-level identifiers from the expression
		topVars := extractTopLevelVars(seg.Text)
		for _, v := range topVars {
			if builtins[v] {
				continue
			}
			if _, ok := commandVars[v]; ok {
				continue
			}
			if _, ok := data[v]; ok {
				continue
			}
			msg := fmt.Sprintf("expression ${%s} references variable %q which is not in data", seg.Text, v)
			if defCell, ok := commandVars[v]; ok {
				msg += fmt.Sprintf(" (provided by command at %s)", defCell)
			}
			issues = append(issues, ValidationIssue{
				Severity: SeverityWarning,
				CellRef:  ref,
				Message:  msg,
			})
		}
	}
	return issues
}

// extractTopLevelVars extracts the top-level identifier names from an expression.
// For "e.Name + company", it returns ["e", "company"].
// For "items[0].Name", it returns ["items"].
// String literals (single and double quoted) are skipped.
func extractTopLevelVars(expression string) []string {
	var vars []string
	seen := map[string]bool{}
	i := 0
	runes := []rune(expression)

	for i < len(runes) {
		ch := runes[i]

		// Skip string literals (single and double quoted)
		if ch == '\'' || ch == '"' {
			quote := ch
			i++ // skip opening quote
			for i < len(runes) && runes[i] != quote {
				if runes[i] == '\\' && i+1 < len(runes) {
					i++ // skip escaped character
				}
				i++
			}
			if i < len(runes) {
				i++ // skip closing quote
			}
			continue
		}

		// Skip backtick strings
		if ch == '`' {
			i++
			for i < len(runes) && runes[i] != '`' {
				i++
			}
			if i < len(runes) {
				i++
			}
			continue
		}

		// Skip non-identifier characters
		if !isIdentStart(ch) {
			i++
			continue
		}

		// Read identifier
		start := i
		for i < len(runes) && isIdentPart(runes[i]) {
			i++
		}
		name := string(runes[start:i])

		// Skip keywords and function calls
		if isExprKeyword(name) {
			continue
		}

		// Check if preceded by '.' (field access, not top-level)
		if start > 0 && runes[start-1] == '.' {
			continue
		}

		if !seen[name] {
			seen[name] = true
			vars = append(vars, name)
		}
	}
	return vars
}

func isIdentStart(r rune) bool {
	return (r >= 'a' && r <= 'z') || (r >= 'A' && r <= 'Z') || r == '_'
}

func isIdentPart(r rune) bool {
	return isIdentStart(r) || (r >= '0' && r <= '9')
}

func isExprKeyword(s string) bool {
	switch s {
	case "true", "false", "nil", "in", "not", "and", "or",
		"matches", "contains", "startsWith", "endsWith",
		"len", "all", "any", "one", "none", "map", "filter",
		"count", "sum", "min", "max", "mean", "first", "last",
		"sort", "sortBy", "groupBy", "toJSON", "fromJSON",
		"trim", "trimPrefix", "trimSuffix", "upper", "lower",
		"split", "join", "repeat", "replace", "indexOf",
		"hasPrefix", "hasSuffix", "sprintf", "int", "float", "string", "bool",
		// Additional expr-lang builtins
		"range", "reverse", "compact", "unique", "flatten",
		"keys", "values", "toPairs", "fromPairs",
		"abs", "ceil", "floor", "round",
		"now", "ago", "duration", "date",
		"type", "get", "take", "reduce",
		"concat", "print", "println":
		return true
	}
	return false
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
			case *RepeatCommand:
				if issue := compileCheck(b.StartRef, "repeat", "count", cmd.Count); issue != nil {
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
