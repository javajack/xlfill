package xlfill

import (
	"runtime"
	"strings"
)

// Mode represents the processing strategy for template filling.
type Mode int

const (
	ModeSequential Mode = iota // default: single-threaded, in-memory
	ModeStreaming              // low-memory output via StreamWriter
	ModeParallel               // concurrent each processing
)

// String returns the mode name.
func (m Mode) String() string {
	switch m {
	case ModeSequential:
		return "sequential"
	case ModeStreaming:
		return "streaming"
	case ModeParallel:
		return "parallel"
	default:
		return "unknown"
	}
}

// ModeSuggestion is the result of analyzing a template to determine the optimal processing mode.
type ModeSuggestion struct {
	Mode        Mode
	Parallelism int      // only set when Mode == ModeParallel
	Reasons     []string // human-readable explanation of why this mode was chosen
}

// SuggestMode analyzes a template and returns the optimal processing mode.
// It inspects command types, formula presence, area structure, and estimated output size
// to decide between sequential, streaming, and parallel modes.
//
// The dataHint parameter provides optional information about the data that will be used:
//   - "itemCount": estimated number of items in the main each loop (int)
//
// Pass nil for dataHint to get a suggestion based on template structure alone.
func SuggestMode(templatePath string, dataHint map[string]any, opts ...Option) (*ModeSuggestion, error) {
	allOpts := append([]Option{WithTemplate(templatePath)}, opts...)
	filler := NewFiller(allOpts...)
	return filler.SuggestMode(dataHint)
}

// SuggestMode analyzes the template and returns the optimal processing mode.
func (f *Filler) SuggestMode(dataHint map[string]any) (*ModeSuggestion, error) {
	tx, err := f.openTemplate()
	if err != nil {
		return nil, err
	}
	defer tx.Close()

	areas, err := f.BuildAreas(tx)
	if err != nil {
		return nil, err
	}

	analysis := analyzeTemplate(areas, tx)

	// Extract hints
	itemCount := 0
	if dataHint != nil {
		if ic, ok := dataHint["itemCount"]; ok {
			itemCount = toInt(ic)
		}
	}

	return decideBestMode(analysis, itemCount), nil
}

// templateAnalysis holds structural properties extracted from a parsed template.
type templateAnalysis struct {
	hasFormulas          bool
	hasImages            bool
	hasHyperlinks        bool
	hasMergeCells        bool
	hasMultiSheet        bool
	hasNestedEach        bool
	hasRepeat            bool
	hasDirectionRight    bool
	hasDataValidation    bool
	hasTable             bool
	hasConditionalFormat bool
	hasChart             bool
	hasSparkline         bool
	hasInclude           bool
	eachCount            int  // number of jx:each commands
	maxEachDepth         int  // deepest nesting level of each commands
	isFixedHeight        bool // all each areas have fixed output height
	sheetCount           int
	estimatedCols        int // max column width across areas
}

// analyzeTemplate walks the parsed area tree and extracts structural properties.
func analyzeTemplate(areas []*Area, tx Transformer) *templateAnalysis {
	a := &templateAnalysis{
		isFixedHeight: true,
		sheetCount:    len(tx.GetSheetNames()),
	}

	// Check for formulas in template
	formulaCells := tx.GetFormulaCells()
	a.hasFormulas = len(formulaCells) > 0

	// Walk command tree
	for _, area := range areas {
		if area.AreaSize.Width > a.estimatedCols {
			a.estimatedCols = area.AreaSize.Width
		}
		analyzeArea(area, a, 0)
	}

	// Check for hyperlinks in cell expressions
	for _, area := range areas {
		scanForHyperlinks(area, tx, a)
	}

	return a
}

// analyzeArea recursively inspects an area and its command bindings.
func analyzeArea(area *Area, a *templateAnalysis, depth int) {
	for _, b := range area.Bindings {
		switch cmd := b.Command.(type) {
		case *EachCommand:
			a.eachCount++
			if depth > 0 {
				a.hasNestedEach = true
			}
			eachDepth := depth + 1
			if eachDepth > a.maxEachDepth {
				a.maxEachDepth = eachDepth
			}
			if cmd.Direction == "RIGHT" {
				a.hasDirectionRight = true
			}
			if cmd.MultiSheet != "" {
				a.hasMultiSheet = true
			}
			if cmd.Area != nil {
				if !isAreaFixedHeight(cmd.Area) {
					a.isFixedHeight = false
				}
				analyzeArea(cmd.Area, a, eachDepth)
			}
		case *RepeatCommand:
			a.hasRepeat = true
			if cmd.Area != nil {
				analyzeArea(cmd.Area, a, depth)
			}
		case *ImageCommand:
			a.hasImages = true
		case *MergeCellsCommand:
			a.hasMergeCells = true
		case *DataValidationCommand:
			a.hasDataValidation = true
			if cmd.Area != nil {
				analyzeArea(cmd.Area, a, depth)
			}
		case *TableCommand:
			a.hasTable = true
			if cmd.Area != nil {
				analyzeArea(cmd.Area, a, depth)
			}
		case *ConditionalFormatCommand:
			a.hasConditionalFormat = true
			if cmd.Area != nil {
				analyzeArea(cmd.Area, a, depth)
			}
		case *GroupCommand:
			if cmd.Area != nil {
				analyzeArea(cmd.Area, a, depth)
			}
		case *ChartCommand:
			a.hasChart = true
			if cmd.Area != nil {
				analyzeArea(cmd.Area, a, depth)
			}
		case *DefinedNameCommand:
			if cmd.Area != nil {
				analyzeArea(cmd.Area, a, depth)
			}
		case *SparklineCommand:
			a.hasSparkline = true
			if cmd.Area != nil {
				analyzeArea(cmd.Area, a, depth)
			}
		case *IncludeCommand:
			a.hasInclude = true
			if cmd.Area != nil {
				analyzeArea(cmd.Area, a, depth)
			}
		case *IfCommand:
			if cmd.IfArea != nil {
				analyzeArea(cmd.IfArea, a, depth)
			}
			if cmd.ElseArea != nil {
				analyzeArea(cmd.ElseArea, a, depth)
			}
		case *GridCommand:
			if cmd.BodyArea != nil {
				analyzeArea(cmd.BodyArea, a, depth)
			}
		default:
			if childArea := getCommandArea(b.Command); childArea != nil {
				analyzeArea(childArea, a, depth)
			}
		}
	}
}

// scanForHyperlinks checks if any cell values or formulas reference hyperlinks.
func scanForHyperlinks(area *Area, tx Transformer, a *templateAnalysis) {
	for row := 0; row < area.AreaSize.Height; row++ {
		for col := 0; col < area.AreaSize.Width; col++ {
			ref := NewCellRef(area.StartCell.Sheet, area.StartCell.Row+row, area.StartCell.Col+col)
			cd := tx.GetCellData(ref)
			if cd == nil {
				continue
			}
			if strVal, ok := cd.Value.(string); ok && strings.Contains(strVal, "hyperlink(") {
				a.hasHyperlinks = true
				return
			}
			if cd.Formula != "" && strings.Contains(strings.ToUpper(cd.Formula), "HYPERLINK(") {
				a.hasHyperlinks = true
				return
			}
		}
	}
}

// decideBestMode picks the optimal mode based on template analysis and data size hints.
func decideBestMode(a *templateAnalysis, itemCount int) *ModeSuggestion {
	s := &ModeSuggestion{Mode: ModeSequential}

	// Check streaming eligibility
	streamingOK := true
	var streamingBlockers []string

	if a.hasFormulas {
		streamingOK = false
		streamingBlockers = append(streamingBlockers, "template has formulas (streaming skips formula remapping)")
	}
	if a.hasImages {
		streamingOK = false
		streamingBlockers = append(streamingBlockers, "template has images (not supported in streaming)")
	}
	if a.hasHyperlinks {
		streamingOK = false
		streamingBlockers = append(streamingBlockers, "template uses hyperlinks (not supported in streaming)")
	}
	if a.hasMultiSheet {
		streamingOK = false
		streamingBlockers = append(streamingBlockers, "template uses multisheet (streaming is single-sheet)")
	}
	if a.hasDirectionRight {
		streamingOK = false
		streamingBlockers = append(streamingBlockers, "template has direction=RIGHT (streaming requires row-sequential writes)")
	}
	if a.hasDataValidation {
		streamingOK = false
		streamingBlockers = append(streamingBlockers, "template has data validation (requires deferred processing)")
	}
	if a.hasTable {
		streamingOK = false
		streamingBlockers = append(streamingBlockers, "template has tables (requires post-processing)")
	}
	if a.hasChart {
		streamingOK = false
		streamingBlockers = append(streamingBlockers, "template has charts (requires deferred processing)")
	}
	if a.hasSparkline {
		streamingOK = false
		streamingBlockers = append(streamingBlockers, "template has sparklines (requires deferred processing)")
	}

	// Check parallel eligibility
	parallelOK := true
	var parallelBlockers []string

	if a.hasMultiSheet {
		parallelOK = false
		parallelBlockers = append(parallelBlockers, "multisheet mode is inherently sequential")
	}
	if !a.isFixedHeight {
		parallelOK = false
		parallelBlockers = append(parallelBlockers, "variable-height areas (nested each/repeat) require sequential offset computation")
	}
	if a.hasDirectionRight {
		parallelOK = false
		parallelBlockers = append(parallelBlockers, "direction=RIGHT needs sequential column tracking")
	}
	if a.eachCount == 0 {
		parallelOK = false
		parallelBlockers = append(parallelBlockers, "no each commands to parallelize")
	}

	// Decision logic:
	// 1. For large estimated output (itemCount > 10000) and streaming-eligible → streaming
	// 2. For moderate output (itemCount > 100) and parallel-eligible → parallel
	// 3. Otherwise → sequential

	// Large output → streaming preferred
	if streamingOK && itemCount >= 10000 {
		s.Mode = ModeStreaming
		s.Reasons = append(s.Reasons, "large dataset (>10K items)")
		s.Reasons = append(s.Reasons, "template is streaming-compatible (no formulas, images, hyperlinks)")
		return s
	}

	// Moderate output → parallel preferred
	cpuCount := runtime.NumCPU()
	if parallelOK && itemCount >= 100 && cpuCount > 1 {
		s.Mode = ModeParallel
		s.Parallelism = cpuCount
		if s.Parallelism > 8 {
			s.Parallelism = 8 // cap at 8 to avoid diminishing returns
		}
		s.Reasons = append(s.Reasons, "moderate dataset (>100 items) with fixed-height areas")
		s.Reasons = append(s.Reasons, "multiple CPU cores available")
		return s
	}

	// Streaming for moderate+ if parallel isn't viable
	if streamingOK && itemCount >= 1000 {
		s.Mode = ModeStreaming
		s.Reasons = append(s.Reasons, "dataset >1K items, streaming reduces memory")
		s.Reasons = append(s.Reasons, "template is streaming-compatible")
		return s
	}

	// Default: sequential
	s.Reasons = append(s.Reasons, "dataset is small or template has constraints")
	if len(streamingBlockers) > 0 {
		s.Reasons = append(s.Reasons, "streaming blocked: "+streamingBlockers[0])
	}
	if len(parallelBlockers) > 0 {
		s.Reasons = append(s.Reasons, "parallel blocked: "+parallelBlockers[0])
	}

	return s
}

// suggestModeFromTransformer analyzes an already-open transformer without closing it.
// Used internally by FillWriter when autoMode is enabled.
func (f *Filler) suggestModeFromTransformer(etx *ExcelizeTransformer, dataHint map[string]any) (*ModeSuggestion, error) {
	areas, err := f.BuildAreas(etx)
	if err != nil {
		return nil, err
	}

	analysis := analyzeTemplate(areas, etx)

	itemCount := 0
	if dataHint != nil {
		if ic, ok := dataHint["itemCount"]; ok {
			itemCount = toInt(ic)
		}
	}

	return decideBestMode(analysis, itemCount), nil
}

// WithAutoMode analyzes the template at fill time and automatically selects
// the optimal processing mode. The dataHint provides sizing information:
//   - "itemCount" (int): estimated number of items in the main each loop
//
// This overrides any explicit WithStreaming or WithParallelism settings.
func WithAutoMode(dataHint map[string]any) Option {
	return func(o *Options) {
		o.autoMode = true
		o.autoModeHint = dataHint
	}
}

// ApplyModeSuggestion configures the Filler options based on a ModeSuggestion.
func (s *ModeSuggestion) Apply(opts ...Option) []Option {
	switch s.Mode {
	case ModeStreaming:
		return append(opts, WithStreaming(true))
	case ModeParallel:
		return append(opts, WithParallelism(s.Parallelism))
	default:
		return opts
	}
}
