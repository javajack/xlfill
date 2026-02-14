package xlfill

import (
	"testing"

	"github.com/stretchr/testify/assert"
)

func TestCellData_Construction(t *testing.T) {
	ref := NewCellRef("Sheet1", 0, 0)
	cd := NewCellData(ref, "Hello", CellString)
	assert.Equal(t, ref, cd.Ref)
	assert.Equal(t, "Hello", cd.Value)
	assert.Equal(t, CellString, cd.Type)
}

func TestCellData_TargetPositions(t *testing.T) {
	cd := NewCellData(NewCellRef("S", 0, 0), nil, CellBlank)
	t1 := NewCellRef("S", 1, 0)
	t2 := NewCellRef("S", 2, 0)

	cd.AddTargetPos(t1)
	cd.AddTargetPos(t2)

	assert.Len(t, cd.TargetPositions, 2)
	assert.Equal(t, t1, cd.TargetPositions[0])
	assert.Equal(t, t2, cd.TargetPositions[1])
}

func TestCellData_AddTargetPosWithArea(t *testing.T) {
	cd := NewCellData(NewCellRef("S", 0, 0), nil, CellBlank)
	target := NewCellRef("S", 5, 0)
	area := NewAreaRef(NewCellRef("S", 5, 0), NewCellRef("S", 10, 5))

	cd.AddTargetPosWithArea(target, area)

	assert.Len(t, cd.TargetPositions, 1)
	assert.Len(t, cd.TargetParentArea, 1)
	assert.Equal(t, target, cd.TargetPositions[0])
	assert.Equal(t, area, cd.TargetParentArea[0])
}

func TestCellData_IsFormulaCell(t *testing.T) {
	formula := NewCellData(NewCellRef("S", 0, 0), nil, CellFormula)
	formula.Formula = "SUM(A1:A10)"
	assert.True(t, formula.IsFormulaCell())

	nonFormula := NewCellData(NewCellRef("S", 0, 0), "text", CellString)
	assert.False(t, nonFormula.IsFormulaCell())

	// Formula set via string even if type is not CellFormula
	hasFormula := NewCellData(NewCellRef("S", 0, 0), nil, CellString)
	hasFormula.Formula = "A1+B1"
	assert.True(t, hasFormula.IsFormulaCell())
}

func TestCellData_Reset(t *testing.T) {
	cd := NewCellData(NewCellRef("S", 0, 0), nil, CellBlank)
	cd.AddTargetPos(NewCellRef("S", 1, 0))
	cd.AddTargetPos(NewCellRef("S", 2, 0))
	cd.EvalResult = "some result"

	cd.Reset()

	assert.Empty(t, cd.TargetPositions)
	assert.Empty(t, cd.TargetParentArea)
	assert.Empty(t, cd.EvalFormulas)
	assert.Nil(t, cd.EvalResult)
}

func TestCellType_String(t *testing.T) {
	assert.Equal(t, "Blank", CellBlank.String())
	assert.Equal(t, "String", CellString.String())
	assert.Equal(t, "Number", CellNumber.String())
	assert.Equal(t, "Boolean", CellBoolean.String())
	assert.Equal(t, "Date", CellDate.String())
	assert.Equal(t, "Formula", CellFormula.String())
	assert.Equal(t, "Error", CellError.String())
}

func TestFormulaStrategy_Constants(t *testing.T) {
	assert.Equal(t, FormulaStrategy(0), FormulaDefault)
	assert.Equal(t, FormulaStrategy(1), FormulaByColumn)
	assert.Equal(t, FormulaStrategy(2), FormulaByRow)
}
