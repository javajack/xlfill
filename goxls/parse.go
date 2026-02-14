package goxls

import (
	"fmt"
	"regexp"
	"strings"
)

const commandPrefix = "jx:"
const paramsPrefix = "jx:params"

// ParsedCommand represents a parsed jx: command from a cell comment.
type ParsedCommand struct {
	Name     string            // command name (e.g., "each", "if", "area")
	Attrs    map[string]string // attributes in order
	LastCell CellRef           // parsed lastCell attribute
	Areas    []AreaRef         // parsed areas attribute (optional)
	CellRef  CellRef           // cell containing this comment
}

// attrPattern matches key="value" pairs, supporting various quote styles.
// Supports: "value", 'value', and Unicode smart quotes from LibreOffice.
var attrPattern = regexp.MustCompile(`(\w+)\s*=\s*["'\x{201C}\x{201D}\x{2018}\x{2019}]([^"'\x{201C}\x{201D}\x{2018}\x{2019}]*)["'\x{201C}\x{201D}\x{2018}\x{2019}]`)

// areasPattern matches the areas=[...] attribute.
var areasPattern = regexp.MustCompile(`areas\s*=\s*\[([^\]]*)\]`)

// areaRefPattern matches cell range references like "A1:C5" or "Sheet1!A1:C5".
var areaRefPattern = regexp.MustCompile(`[A-Za-z0-9_!'.]+:[A-Za-z0-9_!'.]+`)

// ParseComment parses all jx: commands from a cell comment.
// A comment may contain multiple commands (one per line).
func ParseComment(comment string, cellRef CellRef) ([]ParsedCommand, *ParamsData, error) {
	if comment == "" {
		return nil, nil, nil
	}

	lines := splitCommentLines(comment)
	var commands []ParsedCommand
	var params *ParamsData

	for _, line := range lines {
		line = strings.TrimSpace(line)
		if line == "" {
			continue
		}

		if IsParams(line) {
			p, err := ParseParams(line)
			if err != nil {
				return nil, nil, fmt.Errorf("parse params at %s: %w", cellRef, err)
			}
			params = p
			continue
		}

		if !IsCommand(line) {
			continue
		}

		cmd, err := parseCommandLine(line, cellRef)
		if err != nil {
			return nil, nil, fmt.Errorf("parse command at %s: %w", cellRef, err)
		}
		commands = append(commands, cmd)
	}

	return commands, params, nil
}

// splitCommentLines splits a comment into lines, handling both \n and \r\n.
func splitCommentLines(comment string) []string {
	comment = strings.ReplaceAll(comment, "\r\n", "\n")
	comment = strings.ReplaceAll(comment, "\r", "\n")
	return strings.Split(comment, "\n")
}

// IsCommand returns true if the line starts with "jx:" and is not "jx:params".
func IsCommand(line string) bool {
	trimmed := strings.TrimSpace(line)
	return strings.HasPrefix(trimmed, commandPrefix) && !strings.HasPrefix(trimmed, paramsPrefix)
}

// IsParams returns true if the line starts with "jx:params".
func IsParams(line string) bool {
	return strings.HasPrefix(strings.TrimSpace(line), paramsPrefix)
}

// parseCommandLine parses a single command line like:
// jx:each(items="employees" var="e" lastCell="C2")
func parseCommandLine(line string, cellRef CellRef) (ParsedCommand, error) {
	// Extract command name
	nameStart := len(commandPrefix)
	parenIdx := strings.Index(line, "(")
	if parenIdx < 0 {
		return ParsedCommand{}, fmt.Errorf("missing '(' in command: %q", line)
	}
	name := strings.TrimSpace(line[nameStart:parenIdx])

	// Extract attributes string between ( and )
	closeIdx := strings.LastIndex(line, ")")
	if closeIdx < 0 {
		return ParsedCommand{}, fmt.Errorf("missing ')' in command: %q", line)
	}
	attrStr := line[parenIdx+1 : closeIdx]

	// Parse attributes
	attrs := parseAttributes(attrStr)

	// Extract lastCell
	lastCellStr, hasLastCell := attrs["lastCell"]
	if !hasLastCell && name != "params" {
		return ParsedCommand{}, fmt.Errorf("missing lastCell attribute in %s command: %q", name, line)
	}

	var lastCell CellRef
	if hasLastCell {
		var err error
		lastCell, err = ParseCellRef(lastCellStr)
		if err != nil {
			return ParsedCommand{}, fmt.Errorf("invalid lastCell %q: %w", lastCellStr, err)
		}
		// Inherit sheet name from cell if not specified
		if lastCell.Sheet == "" && cellRef.Sheet != "" {
			lastCell.Sheet = cellRef.Sheet
		}
	}

	// Extract areas attribute
	var areas []AreaRef
	areasMatch := areasPattern.FindStringSubmatch(attrStr)
	if len(areasMatch) > 1 {
		areaRefs := areaRefPattern.FindAllString(areasMatch[1], -1)
		for _, ar := range areaRefs {
			areaRef, err := ParseAreaRef(ar)
			if err != nil {
				return ParsedCommand{}, fmt.Errorf("invalid area ref %q: %w", ar, err)
			}
			// Inherit sheet name
			if areaRef.First.Sheet == "" && cellRef.Sheet != "" {
				areaRef.First.Sheet = cellRef.Sheet
				areaRef.Last.Sheet = cellRef.Sheet
			}
			areas = append(areas, areaRef)
		}
	}

	return ParsedCommand{
		Name:     name,
		Attrs:    attrs,
		LastCell: lastCell,
		Areas:    areas,
		CellRef:  cellRef,
	}, nil
}

// parseAttributes extracts key="value" pairs from an attribute string.
func parseAttributes(attrStr string) map[string]string {
	attrs := make(map[string]string)
	matches := attrPattern.FindAllStringSubmatch(attrStr, -1)
	for _, m := range matches {
		if len(m) >= 3 {
			attrs[m[1]] = m[2]
		}
	}
	return attrs
}

// ParamsData holds parsed jx:params attributes.
type ParamsData struct {
	FormulaStrategy FormulaStrategy
	DefaultValue    string
}

// ParseParams parses a jx:params line.
func ParseParams(line string) (*ParamsData, error) {
	parenIdx := strings.Index(line, "(")
	if parenIdx < 0 {
		return &ParamsData{}, nil
	}
	closeIdx := strings.LastIndex(line, ")")
	if closeIdx < 0 {
		return nil, fmt.Errorf("missing ')' in params: %q", line)
	}
	attrStr := line[parenIdx+1 : closeIdx]
	attrs := parseAttributes(attrStr)

	pd := &ParamsData{}

	if dv, ok := attrs["defaultValue"]; ok {
		pd.DefaultValue = dv
	}

	if fs, ok := attrs["formulaStrategy"]; ok {
		switch strings.ToUpper(fs) {
		case "BY_COLUMN":
			pd.FormulaStrategy = FormulaByColumn
		case "BY_ROW":
			pd.FormulaStrategy = FormulaByRow
		default:
			pd.FormulaStrategy = FormulaDefault
		}
	}

	return pd, nil
}
