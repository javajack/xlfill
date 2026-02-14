package xlfill

import (
	"fmt"
	"sort"
	"strings"
)

// Filler orchestrates template processing: parsing, area building, and rendering.
type Filler struct {
	opts     *Options
	registry *CommandRegistry
}

// NewFiller creates a Filler with the given options.
func NewFiller(opts ...Option) *Filler {
	o := defaultOptions()
	for _, opt := range opts {
		opt(o)
	}
	reg := NewCommandRegistry()
	for name, factory := range o.customCommands {
		reg.Register(name, factory)
	}
	return &Filler{opts: o, registry: reg}
}

// BuildAreas parses all commented cells in the transformer and builds the Area/Command hierarchy.
// It finds jx:area commands as root areas, then nests other commands within their containing area.
func (f *Filler) BuildAreas(tx Transformer) ([]*Area, error) {
	commented := tx.GetCommentedCells()
	if len(commented) == 0 {
		return nil, fmt.Errorf("no commented cells found in template")
	}

	type parsedCell struct {
		cellData *CellData
		commands []ParsedCommand
		params   *ParamsData
	}

	var parsed []parsedCell
	for _, cd := range commented {
		cmds, params, _ := ParseComment(cd.Comment, cd.Ref)
		if len(cmds) > 0 || params != nil {
			parsed = append(parsed, parsedCell{cellData: cd, commands: cmds, params: params})
		}
	}

	// Apply params to cell data
	for _, p := range parsed {
		if p.params != nil {
			if p.params.DefaultValue != "" {
				p.cellData.DefaultValue = p.params.DefaultValue
			}
			if p.params.FormulaStrategy != FormulaDefault {
				p.cellData.FormulaStrategy = p.params.FormulaStrategy
			}
		}
	}

	// Find root areas (jx:area commands)
	var rootAreas []*Area

	for _, p := range parsed {
		for _, cmd := range p.commands {
			if cmd.Name != "area" {
				continue
			}
			lastCell := cmd.Attrs["lastCell"]
			if lastCell == "" {
				continue
			}

			startRef := p.cellData.Ref
			endRef, err := resolveLastCell(startRef, lastCell)
			if err != nil {
				return nil, fmt.Errorf("parse area lastCell %q: %w", lastCell, err)
			}

			areaSize := Size{
				Width:  endRef.Col - startRef.Col + 1,
				Height: endRef.Row - startRef.Row + 1,
			}

			area := NewArea(startRef, areaSize, tx)
			rootAreas = append(rootAreas, area)
		}
	}

	if len(rootAreas) == 0 {
		return nil, fmt.Errorf("no jx:area commands found in template")
	}

	// Collect all non-area commands with their parsed info
	type commandInfo struct {
		command  Command
		startRef CellRef
		size     Size
	}
	var allCommands []commandInfo

	for _, p := range parsed {
		for _, cmd := range p.commands {
			if cmd.Name == "area" {
				continue
			}

			command, err := f.registry.Create(cmd.Name, cmd.Attrs)
			if err != nil {
				return nil, fmt.Errorf("create command %q at %s: %w", cmd.Name, p.cellData.Ref, err)
			}
			if command == nil {
				continue // unknown command, silently ignored
			}

			// Parse lastCell to determine command's area size
			lastCell := cmd.Attrs["lastCell"]
			if lastCell == "" {
				continue
			}

			cmdStartRef := p.cellData.Ref
			cmdEndRef, err := resolveLastCell(cmdStartRef, lastCell)
			if err != nil {
				return nil, fmt.Errorf("parse command lastCell %q: %w", lastCell, err)
			}

			cmdSize := Size{
				Width:  cmdEndRef.Col - cmdStartRef.Col + 1,
				Height: cmdEndRef.Row - cmdStartRef.Row + 1,
			}

			// Create the command's inner area and attach it
			innerArea := NewArea(cmdStartRef, cmdSize, tx)
			attachArea(command, innerArea)

			// Handle if command else area (from "areas" attribute)
			if ifCmd, ok := command.(*IfCommand); ok {
				// Use parsed Areas field if available (from areas=[...] syntax)
				if len(cmd.Areas) >= 2 {
					elseAreaRef := cmd.Areas[1]
					elseSize := elseAreaRef.Size()
					ifCmd.ElseArea = NewArea(elseAreaRef.First, elseSize, tx)
				} else if areasAttr := cmd.Attrs["areas"]; areasAttr != "" {
					if err := f.buildIfElseArea(ifCmd, areasAttr, cmdStartRef, tx); err != nil {
						return nil, err
					}
				}
			}

			allCommands = append(allCommands, commandInfo{
				command:  command,
				startRef: cmdStartRef,
				size:     cmdSize,
			})
		}
	}

	// Sort commands by area size (largest first) so parents are processed before children
	sort.Slice(allCommands, func(i, j int) bool {
		areaI := allCommands[i].size.Width * allCommands[i].size.Height
		areaJ := allCommands[j].size.Width * allCommands[j].size.Height
		return areaI > areaJ
	})

	// Build command tree: each command goes into the smallest strictly-larger
	// containing command's area, or into the root area if no parent command contains it.
	// Commands with equal area size are siblings, not parent-child.
	for i, ci := range allCommands {
		ciArea := ci.size.Width * ci.size.Height
		placed := false

		// Find the tightest parent: the smallest command whose area strictly contains ci
		bestParentIdx := -1
		bestParentSize := -1

		for j, cj := range allCommands {
			if i == j {
				continue
			}
			cjArea := cj.size.Width * cj.size.Height
			// Parent must be strictly larger
			if cjArea <= ciArea {
				continue
			}
			parentArea := getCommandArea(cj.command)
			if parentArea == nil {
				continue
			}
			if !parentArea.containsRef(ci.startRef) {
				continue
			}
			// Is this the tightest (smallest valid parent)?
			if bestParentIdx == -1 || cjArea < bestParentSize {
				bestParentIdx = j
				bestParentSize = cjArea
			}
		}

		if bestParentIdx >= 0 {
			parentArea := getCommandArea(allCommands[bestParentIdx].command)
			parentArea.AddCommand(ci.command, ci.startRef, ci.size)
			placed = true
		}

		// If no parent command, add to root area
		if !placed {
			for _, rootArea := range rootAreas {
				if rootArea.containsRef(ci.startRef) {
					rootArea.AddCommand(ci.command, ci.startRef, ci.size)
					break
				}
			}
		}
	}

	// Sort each area's bindings by row then column for deterministic processing
	sortAreaBindings(rootAreas)
	for _, ci := range allCommands {
		if area := getCommandArea(ci.command); area != nil && len(area.Bindings) > 0 {
			sortAreaBindings([]*Area{area})
		}
	}

	// Propagate listeners to all areas (root + command inner areas)
	if len(f.opts.areaListeners) > 0 {
		for _, area := range rootAreas {
			f.propagateListeners(area)
		}
	}

	return rootAreas, nil
}

// propagateListeners sets listeners on an area and all its child command areas recursively.
func (f *Filler) propagateListeners(area *Area) {
	area.Listeners = f.opts.areaListeners
	for _, b := range area.Bindings {
		switch c := b.Command.(type) {
		case *EachCommand:
			if c.Area != nil {
				f.propagateListeners(c.Area)
			}
		case *IfCommand:
			if c.IfArea != nil {
				f.propagateListeners(c.IfArea)
			}
			if c.ElseArea != nil {
				f.propagateListeners(c.ElseArea)
			}
		case *GridCommand:
			if c.BodyArea != nil {
				f.propagateListeners(c.BodyArea)
			}
		case *UpdateCellCommand:
			if c.Area != nil {
				f.propagateListeners(c.Area)
			}
		case *AutoRowHeightCommand:
			if c.Area != nil {
				f.propagateListeners(c.Area)
			}
		}
	}
}

// getCommandArea returns the inner area of a command, or nil if the command type has no area.
func getCommandArea(cmd Command) *Area {
	switch c := cmd.(type) {
	case *EachCommand:
		return c.Area
	case *IfCommand:
		return c.IfArea
	case *UpdateCellCommand:
		return c.Area
	case *GridCommand:
		return c.BodyArea
	case *AutoRowHeightCommand:
		return c.Area
	}
	return nil
}

// sortAreaBindings sorts bindings in each area by row then column.
func sortAreaBindings(areas []*Area) {
	for _, area := range areas {
		sort.Slice(area.Bindings, func(i, j int) bool {
			bi, bj := area.Bindings[i], area.Bindings[j]
			if bi.StartRef.Row != bj.StartRef.Row {
				return bi.StartRef.Row < bj.StartRef.Row
			}
			return bi.StartRef.Col < bj.StartRef.Col
		})
	}
}

// buildIfElseArea parses the "areas" attribute to set up the else area for an IfCommand.
// Format: areas=["A2:C2", "A3:C3"] — first is if area (already set), second is else area.
func (f *Filler) buildIfElseArea(ifCmd *IfCommand, areasAttr string, cmdStart CellRef, tx Transformer) error {
	// Parse areas: ["ref1", "ref2"] or "ref1, ref2"
	areasAttr = strings.Trim(areasAttr, "[]")
	parts := strings.Split(areasAttr, ",")
	if len(parts) < 2 {
		return nil // no else area
	}

	elseRef := strings.Trim(strings.TrimSpace(parts[1]), "\"' ")
	if elseRef == "" {
		return nil
	}

	// Parse the else area reference
	areaRef, err := ParseAreaRef(cmdStart.Sheet + "!" + elseRef)
	if err != nil {
		// Try without sheet
		areaRef, err = ParseAreaRef(elseRef)
		if err != nil {
			return fmt.Errorf("parse if else area %q: %w", elseRef, err)
		}
		if areaRef.First.Sheet == "" {
			areaRef.First.Sheet = cmdStart.Sheet
			areaRef.Last.Sheet = cmdStart.Sheet
		}
	}

	elseSize := areaRef.Size()
	ifCmd.ElseArea = NewArea(areaRef.First, elseSize, tx)
	return nil
}

// attachArea attaches an inner area to a command based on its type.
func attachArea(cmd Command, area *Area) {
	switch c := cmd.(type) {
	case *EachCommand:
		c.Area = area
	case *IfCommand:
		c.IfArea = area
	case *UpdateCellCommand:
		c.Area = area
	case *GridCommand:
		c.BodyArea = area
	case *AutoRowHeightCommand:
		c.Area = area
	}
}

// containsRef checks if a cell reference is within this area.
func (a *Area) containsRef(ref CellRef) bool {
	if ref.Sheet != a.StartCell.Sheet {
		return false
	}
	return ref.Row >= a.StartCell.Row &&
		ref.Row < a.StartCell.Row+a.AreaSize.Height &&
		ref.Col >= a.StartCell.Col &&
		ref.Col < a.StartCell.Col+a.AreaSize.Width
}

// resolveLastCell resolves a lastCell reference relative to a start cell.
func resolveLastCell(start CellRef, lastCell string) (CellRef, error) {
	// If lastCell contains "!", it has its own sheet
	if strings.Contains(lastCell, "!") {
		return ParseCellRef(lastCell)
	}
	ref, err := ParseCellRef(lastCell)
	if err != nil {
		return CellRef{}, err
	}
	ref.Sheet = start.Sheet
	return ref, nil
}
