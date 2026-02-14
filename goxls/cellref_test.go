package goxls

import (
	"testing"

	"github.com/stretchr/testify/assert"
	"github.com/stretchr/testify/require"
)

// --- CellRef Tests ---

func TestParseCellRef_SimpleCell(t *testing.T) {
	ref, err := ParseCellRef("A1")
	require.NoError(t, err)
	assert.Equal(t, "", ref.Sheet)
	assert.Equal(t, 0, ref.Row)
	assert.Equal(t, 0, ref.Col)
}

func TestParseCellRef_WithSheet(t *testing.T) {
	ref, err := ParseCellRef("Sheet1!B5")
	require.NoError(t, err)
	assert.Equal(t, "Sheet1", ref.Sheet)
	assert.Equal(t, 4, ref.Row) // 0-based
	assert.Equal(t, 1, ref.Col)
}

func TestParseCellRef_AbsoluteRef(t *testing.T) {
	ref, err := ParseCellRef("$A$1")
	require.NoError(t, err)
	assert.Equal(t, 0, ref.Row)
	assert.Equal(t, 0, ref.Col)
}

func TestParseCellRef_MultiLetterCol(t *testing.T) {
	ref, err := ParseCellRef("AA1")
	require.NoError(t, err)
	assert.Equal(t, 0, ref.Row)
	assert.Equal(t, 26, ref.Col)
}

func TestParseCellRef_LargeCol(t *testing.T) {
	ref, err := ParseCellRef("AZ10")
	require.NoError(t, err)
	assert.Equal(t, 9, ref.Row)
	assert.Equal(t, 51, ref.Col) // AZ = 26+25 = 51
}

func TestParseCellRef_Invalid_Empty(t *testing.T) {
	_, err := ParseCellRef("")
	assert.Error(t, err)
}

func TestParseCellRef_Invalid_NoRow(t *testing.T) {
	_, err := ParseCellRef("A")
	assert.Error(t, err)
}

func TestParseCellRef_Invalid_NoCol(t *testing.T) {
	_, err := ParseCellRef("123")
	assert.Error(t, err)
}

func TestParseCellRef_QuotedSheet(t *testing.T) {
	ref, err := ParseCellRef("'My Sheet'!A1")
	require.NoError(t, err)
	assert.Equal(t, "My Sheet", ref.Sheet)
	assert.Equal(t, 0, ref.Row)
	assert.Equal(t, 0, ref.Col)
}

func TestCellRef_String(t *testing.T) {
	ref := NewCellRef("Sheet1", 4, 1)
	assert.Equal(t, "Sheet1!B5", ref.String())
}

func TestCellRef_String_NoSheet(t *testing.T) {
	ref := NewCellRef("", 0, 0)
	assert.Equal(t, "A1", ref.String())
}

func TestCellRef_CellName(t *testing.T) {
	ref := NewCellRef("Sheet1", 9, 2)
	assert.Equal(t, "C10", ref.CellName())
}

func TestCellRef_Roundtrip(t *testing.T) {
	cases := []string{"A1", "Z99", "AA1", "Sheet1!B5"}
	for _, tc := range cases {
		ref, err := ParseCellRef(tc)
		require.NoError(t, err, "parse %q", tc)
		assert.Equal(t, tc, ref.String(), "roundtrip %q", tc)
	}
}

// --- ColToName / NameToCol Tests ---

func TestColToName(t *testing.T) {
	tests := map[int]string{
		0:   "A",
		1:   "B",
		25:  "Z",
		26:  "AA",
		27:  "AB",
		51:  "AZ",
		52:  "BA",
		701: "ZZ",
		702: "AAA",
	}
	for col, expected := range tests {
		assert.Equal(t, expected, ColToName(col), "col %d", col)
	}
}

func TestNameToCol(t *testing.T) {
	tests := map[string]int{
		"A":   0,
		"B":   1,
		"Z":   25,
		"AA":  26,
		"AB":  27,
		"AZ":  51,
		"BA":  52,
		"ZZ":  701,
		"AAA": 702,
	}
	for name, expected := range tests {
		col, err := NameToCol(name)
		require.NoError(t, err, "name %q", name)
		assert.Equal(t, expected, col, "name %q", name)
	}
}

func TestNameToCol_CaseInsensitive(t *testing.T) {
	col, err := NameToCol("aa")
	require.NoError(t, err)
	assert.Equal(t, 26, col)
}

func TestNameToCol_Invalid(t *testing.T) {
	_, err := NameToCol("")
	assert.Error(t, err)
	_, err = NameToCol("1A")
	assert.Error(t, err)
}

func TestColToName_NameToCol_Roundtrip(t *testing.T) {
	for i := 0; i < 1000; i++ {
		name := ColToName(i)
		col, err := NameToCol(name)
		require.NoError(t, err)
		assert.Equal(t, i, col, "roundtrip col %d → %q → %d", i, name, col)
	}
}

// --- AreaRef Tests ---

func TestParseAreaRef(t *testing.T) {
	ar, err := ParseAreaRef("A1:C5")
	require.NoError(t, err)
	assert.Equal(t, 0, ar.First.Row)
	assert.Equal(t, 0, ar.First.Col)
	assert.Equal(t, 4, ar.Last.Row)
	assert.Equal(t, 2, ar.Last.Col)
}

func TestParseAreaRef_WithSheet(t *testing.T) {
	ar, err := ParseAreaRef("Sheet1!A1:C5")
	require.NoError(t, err)
	assert.Equal(t, "Sheet1", ar.First.Sheet)
	assert.Equal(t, "Sheet1", ar.Last.Sheet)
}

func TestParseAreaRef_Invalid(t *testing.T) {
	_, err := ParseAreaRef("A1")
	assert.Error(t, err)
}

func TestAreaRef_Size(t *testing.T) {
	ar, _ := ParseAreaRef("B2:D6")
	s := ar.Size()
	assert.Equal(t, 3, s.Width)  // B,C,D
	assert.Equal(t, 5, s.Height) // rows 2-6
}

func TestAreaRef_String(t *testing.T) {
	ar := NewAreaRef(NewCellRef("Sheet1", 0, 0), NewCellRef("Sheet1", 4, 2))
	assert.Equal(t, "Sheet1!A1:C5", ar.String())
}

func TestAreaRef_String_NoSheet(t *testing.T) {
	ar := NewAreaRef(NewCellRef("", 0, 0), NewCellRef("", 4, 2))
	assert.Equal(t, "A1:C5", ar.String())
}

// --- AreaRef.Contains Tests (parity with JXLS AreaRefContainsTest) ---

func TestAreaRef_Contains_Inside(t *testing.T) {
	ar := NewAreaRef(NewCellRef("S", 1, 1), NewCellRef("S", 5, 5))
	assert.True(t, ar.Contains(NewCellRef("S", 3, 3)))
}

func TestAreaRef_Contains_TopLeft(t *testing.T) {
	ar := NewAreaRef(NewCellRef("S", 1, 1), NewCellRef("S", 5, 5))
	assert.True(t, ar.Contains(NewCellRef("S", 1, 1)))
}

func TestAreaRef_Contains_TopRight(t *testing.T) {
	ar := NewAreaRef(NewCellRef("S", 1, 1), NewCellRef("S", 5, 5))
	assert.True(t, ar.Contains(NewCellRef("S", 1, 5)))
}

func TestAreaRef_Contains_BottomLeft(t *testing.T) {
	ar := NewAreaRef(NewCellRef("S", 1, 1), NewCellRef("S", 5, 5))
	assert.True(t, ar.Contains(NewCellRef("S", 5, 1)))
}

func TestAreaRef_Contains_BottomRight(t *testing.T) {
	ar := NewAreaRef(NewCellRef("S", 1, 1), NewCellRef("S", 5, 5))
	assert.True(t, ar.Contains(NewCellRef("S", 5, 5)))
}

func TestAreaRef_Contains_Outside_Left(t *testing.T) {
	ar := NewAreaRef(NewCellRef("S", 1, 1), NewCellRef("S", 5, 5))
	assert.False(t, ar.Contains(NewCellRef("S", 3, 0)))
}

func TestAreaRef_Contains_Outside_Right(t *testing.T) {
	ar := NewAreaRef(NewCellRef("S", 1, 1), NewCellRef("S", 5, 5))
	assert.False(t, ar.Contains(NewCellRef("S", 3, 6)))
}

func TestAreaRef_Contains_Outside_Above(t *testing.T) {
	ar := NewAreaRef(NewCellRef("S", 1, 1), NewCellRef("S", 5, 5))
	assert.False(t, ar.Contains(NewCellRef("S", 0, 3)))
}

func TestAreaRef_Contains_Outside_Below(t *testing.T) {
	ar := NewAreaRef(NewCellRef("S", 1, 1), NewCellRef("S", 5, 5))
	assert.False(t, ar.Contains(NewCellRef("S", 6, 3)))
}

func TestAreaRef_Contains_DifferentSheet(t *testing.T) {
	ar := NewAreaRef(NewCellRef("Sheet1", 1, 1), NewCellRef("Sheet1", 5, 5))
	assert.False(t, ar.Contains(NewCellRef("Sheet2", 3, 3)))
}

func TestAreaRef_Contains_EmptySheet(t *testing.T) {
	// Area with no sheet matches any sheet
	ar := NewAreaRef(NewCellRef("", 1, 1), NewCellRef("", 5, 5))
	assert.True(t, ar.Contains(NewCellRef("AnySheet", 3, 3)))
}

// --- Size Tests (parity with JXLS SizeTest) ---

func TestSize_String(t *testing.T) {
	assert.Equal(t, "(3x5)", Size{Width: 3, Height: 5}.String())
}

func TestSize_Add(t *testing.T) {
	s := Size{Width: 2, Height: 3}.Add(Size{Width: 1, Height: 4})
	assert.Equal(t, Size{Width: 3, Height: 7}, s)
}

func TestSize_Minus(t *testing.T) {
	s := Size{Width: 5, Height: 5}.Minus(Size{Width: 2, Height: 3})
	assert.Equal(t, Size{Width: 3, Height: 2}, s)
}

func TestSize_Zero(t *testing.T) {
	assert.Equal(t, 0, ZeroSize.Width)
	assert.Equal(t, 0, ZeroSize.Height)
}
