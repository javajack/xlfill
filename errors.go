package xlfill

import (
	"fmt"
	"strings"
	"sync"
)

// ErrorKind categorizes xlfill errors for programmatic handling.
type ErrorKind int

const (
	ErrTemplate ErrorKind = iota // template structure errors (bad jx: syntax, missing area)
	ErrData                      // data/expression errors (missing key, wrong type)
	ErrRuntime                   // runtime errors (file I/O, excelize failures)
)

// String returns the error kind name.
func (k ErrorKind) String() string {
	switch k {
	case ErrTemplate:
		return "template"
	case ErrData:
		return "data"
	case ErrRuntime:
		return "runtime"
	default:
		return "unknown"
	}
}

// XLFillError is the structured error type for xlfill operations.
// It satisfies the error and Unwrap interfaces for use with errors.Is/As.
type XLFillError struct {
	Kind    ErrorKind
	Cell    CellRef
	Command string // command name, if applicable
	Message string
	Err     error // wrapped cause
}

// Error implements the error interface.
func (e *XLFillError) Error() string {
	var b strings.Builder
	b.WriteString("[xlfill:")
	b.WriteString(e.Kind.String())
	b.WriteString("] ")
	if e.Cell != (CellRef{}) {
		b.WriteString(e.Cell.String())
		b.WriteString(": ")
	}
	if e.Command != "" {
		b.WriteString(e.Command)
		b.WriteString(": ")
	}
	b.WriteString(e.Message)
	if e.Err != nil {
		b.WriteString(": ")
		b.WriteString(e.Err.Error())
	}
	return b.String()
}

// Unwrap returns the underlying error.
func (e *XLFillError) Unwrap() error {
	return e.Err
}

// NewTemplateError creates a template structure error.
func NewTemplateError(cell CellRef, command, message string, err error) *XLFillError {
	return &XLFillError{Kind: ErrTemplate, Cell: cell, Command: command, Message: message, Err: err}
}

// NewDataError creates a data/expression error.
func NewDataError(cell CellRef, message string, err error) *XLFillError {
	return &XLFillError{Kind: ErrData, Cell: cell, Message: message, Err: err}
}

// NewRuntimeError creates a runtime/IO error.
func NewRuntimeError(message string, err error) *XLFillError {
	return &XLFillError{Kind: ErrRuntime, Message: message, Err: err}
}

// Warning represents a non-fatal issue detected during template processing.
type Warning struct {
	Cell    CellRef
	Message string
}

// String formats the warning.
func (w Warning) String() string {
	if w.Cell != (CellRef{}) {
		return fmt.Sprintf("[WARN] %s: %s", w.Cell, w.Message)
	}
	return fmt.Sprintf("[WARN] %s", w.Message)
}

// WarningCollector accumulates warnings during processing.
type WarningCollector struct {
	mu       sync.Mutex
	warnings []Warning
}

// Add records a warning.
func (wc *WarningCollector) Add(cell CellRef, message string) {
	wc.mu.Lock()
	wc.warnings = append(wc.warnings, Warning{Cell: cell, Message: message})
	wc.mu.Unlock()
}

// Warnings returns all collected warnings.
func (wc *WarningCollector) Warnings() []Warning {
	wc.mu.Lock()
	defer wc.mu.Unlock()
	return wc.warnings
}

// Reset clears all warnings.
func (wc *WarningCollector) Reset() {
	wc.mu.Lock()
	wc.warnings = wc.warnings[:0]
	wc.mu.Unlock()
}

// levenshtein computes the edit distance between two strings.
func levenshtein(a, b string) int {
	la, lb := len(a), len(b)
	if la == 0 {
		return lb
	}
	if lb == 0 {
		return la
	}
	prev := make([]int, lb+1)
	curr := make([]int, lb+1)
	for j := 0; j <= lb; j++ {
		prev[j] = j
	}
	for i := 1; i <= la; i++ {
		curr[0] = i
		for j := 1; j <= lb; j++ {
			cost := 1
			if a[i-1] == b[j-1] {
				cost = 0
			}
			curr[j] = min(curr[j-1]+1, min(prev[j]+1, prev[j-1]+cost))
		}
		prev, curr = curr, prev
	}
	return prev[lb]
}

// suggestCommand returns a "did you mean ...?" hint if name is close to a known command.
func suggestCommand(name string, knownNames []string) string {
	bestDist := 3 // max edit distance to suggest
	bestName := ""
	for _, known := range knownNames {
		d := levenshtein(strings.ToLower(name), strings.ToLower(known))
		if d < bestDist {
			bestDist = d
			bestName = known
		}
	}
	if bestName != "" {
		return fmt.Sprintf(" (did you mean %q?)", bestName)
	}
	return ""
}
