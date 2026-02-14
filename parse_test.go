package xlfill

import (
	"testing"

	"github.com/stretchr/testify/assert"
	"github.com/stretchr/testify/require"
)

func cell(sheet string, row, col int) CellRef {
	return NewCellRef(sheet, row, col)
}

func TestParseComment_SimpleEach(t *testing.T) {
	cmds, _, err := ParseComment(`jx:each(items="employees" var="e" lastCell="C2")`, cell("S", 1, 0))
	require.NoError(t, err)
	require.Len(t, cmds, 1)

	c := cmds[0]
	assert.Equal(t, "each", c.Name)
	assert.Equal(t, "employees", c.Attrs["items"])
	assert.Equal(t, "e", c.Attrs["var"])
	assert.Equal(t, "S", c.LastCell.Sheet)
	assert.Equal(t, 1, c.LastCell.Row) // C2 → row 1
	assert.Equal(t, 2, c.LastCell.Col) // C → col 2
}

func TestParseComment_IfWithAreas(t *testing.T) {
	cmds, _, err := ParseComment(
		`jx:if(condition="x>1" lastCell="C2" areas=["A2:C2","A3:C3"])`,
		cell("S", 1, 0),
	)
	require.NoError(t, err)
	require.Len(t, cmds, 1)

	c := cmds[0]
	assert.Equal(t, "if", c.Name)
	assert.Equal(t, "x>1", c.Attrs["condition"])
	require.Len(t, c.Areas, 2)
	assert.Equal(t, "S!A2:C2", c.Areas[0].String())
	assert.Equal(t, "S!A3:C3", c.Areas[1].String())
}

func TestParseComment_Area(t *testing.T) {
	cmds, _, err := ParseComment(`jx:area(lastCell="D10")`, cell("Sheet1", 0, 0))
	require.NoError(t, err)
	require.Len(t, cmds, 1)

	assert.Equal(t, "area", cmds[0].Name)
	assert.Equal(t, "Sheet1", cmds[0].LastCell.Sheet)
	assert.Equal(t, 9, cmds[0].LastCell.Row) // D10 → row 9
	assert.Equal(t, 3, cmds[0].LastCell.Col) // D → col 3
}

func TestParseComment_Grid(t *testing.T) {
	cmds, _, err := ParseComment(
		`jx:grid(headers="h" data="d" areas=["A1:A1","A2:A2"] lastCell="A2")`,
		cell("S", 0, 0),
	)
	require.NoError(t, err)
	require.Len(t, cmds, 1)

	c := cmds[0]
	assert.Equal(t, "grid", c.Name)
	assert.Equal(t, "h", c.Attrs["headers"])
	assert.Equal(t, "d", c.Attrs["data"])
	require.Len(t, c.Areas, 2)
}

func TestParseComment_Image(t *testing.T) {
	cmds, _, err := ParseComment(
		`jx:image(src="img" imageType="PNG" lastCell="A2")`,
		cell("S", 0, 0),
	)
	require.NoError(t, err)
	require.Len(t, cmds, 1)

	assert.Equal(t, "image", cmds[0].Name)
	assert.Equal(t, "img", cmds[0].Attrs["src"])
	assert.Equal(t, "PNG", cmds[0].Attrs["imageType"])
}

func TestParseComment_MergeCells(t *testing.T) {
	cmds, _, err := ParseComment(
		`jx:mergeCells(lastCell="D2" cols="4" rows="2")`,
		cell("S", 0, 0),
	)
	require.NoError(t, err)
	require.Len(t, cmds, 1)

	assert.Equal(t, "mergeCells", cmds[0].Name)
	assert.Equal(t, "4", cmds[0].Attrs["cols"])
	assert.Equal(t, "2", cmds[0].Attrs["rows"])
}

func TestParseComment_UpdateCell(t *testing.T) {
	cmds, _, err := ParseComment(
		`jx:updateCell(lastCell="E4" updater="myUpdater")`,
		cell("S", 0, 0),
	)
	require.NoError(t, err)
	require.Len(t, cmds, 1)

	assert.Equal(t, "updateCell", cmds[0].Name)
	assert.Equal(t, "myUpdater", cmds[0].Attrs["updater"])
}

func TestParseComment_MultipleCommands(t *testing.T) {
	comment := "jx:area(lastCell=\"C3\")\njx:each(items=\"list\" var=\"e\" lastCell=\"C2\")"
	cmds, _, err := ParseComment(comment, cell("S", 0, 0))
	require.NoError(t, err)
	require.Len(t, cmds, 2)

	assert.Equal(t, "area", cmds[0].Name)
	assert.Equal(t, "each", cmds[1].Name)
}

func TestParseComment_EachAllAttrs(t *testing.T) {
	cmds, _, err := ParseComment(
		`jx:each(items="list" var="e" varIndex="idx" direction="RIGHT" select="e.Active" groupBy="e.Dept" groupOrder="ASC" orderBy="e.Name ASC" multisheet="sheets" lastCell="C2")`,
		cell("S", 0, 0),
	)
	require.NoError(t, err)
	require.Len(t, cmds, 1)

	c := cmds[0]
	assert.Equal(t, "list", c.Attrs["items"])
	assert.Equal(t, "e", c.Attrs["var"])
	assert.Equal(t, "idx", c.Attrs["varIndex"])
	assert.Equal(t, "RIGHT", c.Attrs["direction"])
	assert.Equal(t, "e.Active", c.Attrs["select"])
	assert.Equal(t, "e.Dept", c.Attrs["groupBy"])
	assert.Equal(t, "ASC", c.Attrs["groupOrder"])
	assert.Equal(t, "e.Name ASC", c.Attrs["orderBy"])
	assert.Equal(t, "sheets", c.Attrs["multisheet"])
}

func TestParseComment_WithCommas(t *testing.T) {
	// Optional commas between attributes (JXLS supports this)
	cmds, _, err := ParseComment(
		`jx:each(items="list", var="e", lastCell="C2")`,
		cell("S", 0, 0),
	)
	require.NoError(t, err)
	require.Len(t, cmds, 1)
	assert.Equal(t, "list", cmds[0].Attrs["items"])
	assert.Equal(t, "e", cmds[0].Attrs["var"])
}

func TestParseComment_WhitespaceVariants(t *testing.T) {
	cmds, _, err := ParseComment(
		`jx:each( items = "list"   var = "e"   lastCell = "C2" )`,
		cell("S", 0, 0),
	)
	require.NoError(t, err)
	require.Len(t, cmds, 1)
	assert.Equal(t, "list", cmds[0].Attrs["items"])
}

func TestParseComment_SheetInLastCell(t *testing.T) {
	cmds, _, err := ParseComment(
		`jx:area(lastCell="Sheet2!A5")`,
		cell("Sheet1", 0, 0),
	)
	require.NoError(t, err)
	require.Len(t, cmds, 1)
	assert.Equal(t, "Sheet2", cmds[0].LastCell.Sheet)
	assert.Equal(t, 4, cmds[0].LastCell.Row) // A5 → row 4
}

func TestParseComment_InvalidCommand_MissingLastCell(t *testing.T) {
	_, _, err := ParseComment(
		`jx:each(items="list" var="e")`,
		cell("S", 0, 0),
	)
	assert.Error(t, err)
}

func TestParseComment_EmptyComment(t *testing.T) {
	cmds, params, err := ParseComment("", cell("S", 0, 0))
	require.NoError(t, err)
	assert.Empty(t, cmds)
	assert.Nil(t, params)
}

func TestParseComment_NonJxComment(t *testing.T) {
	cmds, _, err := ParseComment("This is a regular note", cell("S", 0, 0))
	require.NoError(t, err)
	assert.Empty(t, cmds)
}

func TestIsCommand(t *testing.T) {
	assert.True(t, IsCommand(`jx:each(items="x" lastCell="A1")`))
	assert.True(t, IsCommand(`jx:area(lastCell="A1")`))
	assert.True(t, IsCommand(`jx:if(condition="true" lastCell="A1")`))
	assert.False(t, IsCommand(`jx:params(defaultValue="1")`))
	assert.False(t, IsCommand("This is a note"))
	assert.False(t, IsCommand(""))
}

func TestIsParams(t *testing.T) {
	assert.True(t, IsParams(`jx:params(defaultValue="1")`))
	assert.True(t, IsParams(`jx:params(formulaStrategy="BY_COLUMN")`))
	assert.False(t, IsParams(`jx:each(items="x" lastCell="A1")`))
	assert.False(t, IsParams(""))
}

func TestParseParams_DefaultValue(t *testing.T) {
	_, params, err := ParseComment(`jx:params(defaultValue="1")`, cell("S", 0, 0))
	require.NoError(t, err)
	require.NotNil(t, params)
	assert.Equal(t, "1", params.DefaultValue)
	assert.Equal(t, FormulaDefault, params.FormulaStrategy)
}

func TestParseParams_FormulaStrategy(t *testing.T) {
	_, params, err := ParseComment(`jx:params(formulaStrategy="BY_COLUMN")`, cell("S", 0, 0))
	require.NoError(t, err)
	require.NotNil(t, params)
	assert.Equal(t, FormulaByColumn, params.FormulaStrategy)
}

func TestParseParams_Both(t *testing.T) {
	_, params, err := ParseComment(
		`jx:params(defaultValue="0" formulaStrategy="BY_ROW")`,
		cell("S", 0, 0),
	)
	require.NoError(t, err)
	require.NotNil(t, params)
	assert.Equal(t, "0", params.DefaultValue)
	assert.Equal(t, FormulaByRow, params.FormulaStrategy)
}

func TestParseComment_CommandAndParams(t *testing.T) {
	comment := "jx:each(items=\"list\" var=\"e\" lastCell=\"C2\")\njx:params(defaultValue=\"1\")"
	cmds, params, err := ParseComment(comment, cell("S", 0, 0))
	require.NoError(t, err)
	require.Len(t, cmds, 1)
	assert.Equal(t, "each", cmds[0].Name)
	require.NotNil(t, params)
	assert.Equal(t, "1", params.DefaultValue)
}

func TestParseComment_SingleQuotes(t *testing.T) {
	cmds, _, err := ParseComment(
		`jx:each(items='list' var='e' lastCell='C2')`,
		cell("S", 0, 0),
	)
	require.NoError(t, err)
	require.Len(t, cmds, 1)
	assert.Equal(t, "list", cmds[0].Attrs["items"])
}

func TestParseComment_MultiLine(t *testing.T) {
	comment := "jx:area(lastCell=\"C5\")\r\njx:each(items=\"employees\" var=\"e\" lastCell=\"C2\")"
	cmds, _, err := ParseComment(comment, cell("S", 0, 0))
	require.NoError(t, err)
	require.Len(t, cmds, 2)
}
