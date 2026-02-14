package goxls

// CellType represents the type of data in a cell.
type CellType int

const (
	CellBlank CellType = iota
	CellString
	CellNumber
	CellBoolean
	CellDate
	CellFormula
	CellError
)

// String returns a human-readable name for the CellType.
func (ct CellType) String() string {
	switch ct {
	case CellBlank:
		return "Blank"
	case CellString:
		return "String"
	case CellNumber:
		return "Number"
	case CellBoolean:
		return "Boolean"
	case CellDate:
		return "Date"
	case CellFormula:
		return "Formula"
	case CellError:
		return "Error"
	default:
		return "Unknown"
	}
}

// FormulaStrategy controls how formula cell references are resolved
// during formula processing after template expansion.
type FormulaStrategy int

const (
	FormulaDefault  FormulaStrategy = iota // references expand to all target cells
	FormulaByColumn                        // only reference cells in the same column
	FormulaByRow                           // only reference cells in the same row
)

// CellData holds all information about a single cell in the template.
type CellData struct {
	Ref             CellRef         // cell position
	Value           any             // cell value
	Type            CellType        // value type
	Comment         string          // cell comment/note text
	Formula         string          // Excel formula (without leading =)
	EvalResult      any             // result of expression evaluation
	TargetCellType  CellType        // type to use when writing to target
	FormulaStrategy FormulaStrategy // formula expansion strategy (from jx:params)
	DefaultValue    string          // default value for removed formula refs (from jx:params)

	// Tracking for formula processing
	TargetPositions  []CellRef  // where this cell was copied to during transformation
	TargetParentArea []AreaRef  // parent area of each target position
	EvalFormulas     []string   // evaluated formulas for each target position

	// Style preservation
	StyleID int // cached style ID for restoring after value write
}

// NewCellData creates a CellData with a reference, value, and type.
func NewCellData(ref CellRef, value any, cellType CellType) *CellData {
	return &CellData{
		Ref:   ref,
		Value: value,
		Type:  cellType,
	}
}

// AddTargetPos records that this cell was copied to the given target position.
func (cd *CellData) AddTargetPos(ref CellRef) {
	cd.TargetPositions = append(cd.TargetPositions, ref)
}

// AddTargetPosWithArea records a target position with its parent area.
func (cd *CellData) AddTargetPosWithArea(ref CellRef, area AreaRef) {
	cd.TargetPositions = append(cd.TargetPositions, ref)
	cd.TargetParentArea = append(cd.TargetParentArea, area)
}

// IsFormulaCell returns true if this cell contains a formula.
func (cd *CellData) IsFormulaCell() bool {
	return cd.Type == CellFormula || cd.Formula != ""
}

// Reset clears target tracking data for reuse.
func (cd *CellData) Reset() {
	cd.TargetPositions = cd.TargetPositions[:0]
	cd.TargetParentArea = cd.TargetParentArea[:0]
	cd.EvalFormulas = cd.EvalFormulas[:0]
	cd.EvalResult = nil
}
