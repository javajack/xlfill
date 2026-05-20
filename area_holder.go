package xlfill

// AreaHolder is implemented by any command that owns an inner Area which the
// engine needs to attach during BuildAreas and walk during listener
// propagation. Adding a new command with its own area is now a one-line
// interface implementation rather than a 4-place type-switch update.
type AreaHolder interface {
	// GetArea returns the command's primary inner area (or nil if not set).
	GetArea() *Area
	// SetArea installs the command's primary inner area.
	SetArea(*Area)
}

// MultiAreaHolder is implemented by commands that own more than one inner
// area (currently only jx:if with its if/else branches). The Areas slice
// MUST include the primary area returned by GetArea, plus any additional
// areas, in walk order.
type MultiAreaHolder interface {
	AreaHolder
	Areas() []*Area
}

// commandAreas returns every inner area of cmd in walk order.
// Returns nil if cmd holds no area at all.
func commandAreas(cmd Command) []*Area {
	if mh, ok := cmd.(MultiAreaHolder); ok {
		return mh.Areas()
	}
	if h, ok := cmd.(AreaHolder); ok {
		if a := h.GetArea(); a != nil {
			return []*Area{a}
		}
	}
	return nil
}
