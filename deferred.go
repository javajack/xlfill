package xlfill

import (
	"fmt"
	"sync"
)

// MaxDeferredActions limits the number of deferred actions to prevent resource exhaustion.
const MaxDeferredActions = 10_000

// DeferredAction represents an action to be executed after all areas have been processed.
// Commands like jx:table, jx:chart register deferred actions during ApplyAt because they
// need to know the FINAL output range after jx:each expansion.
type DeferredAction struct {
	// Name identifies the deferred action (e.g., "table", "chart").
	Name string

	// Sheet is the target sheet name.
	Sheet string

	// StartRow, StartCol define the top-left of the target range (0-based).
	StartRow int
	StartCol int

	// EndRow, EndCol define the bottom-right of the target range (0-based).
	EndRow int
	EndCol int

	// Execute is called after all areas are processed.
	// The ExcelizeTransformer is passed for direct excelize API calls.
	Execute func(tx *ExcelizeTransformer) error
}

// DeferredRegistry collects deferred actions during template processing.
// It is safe for concurrent use.
type DeferredRegistry struct {
	mu      sync.Mutex
	actions []DeferredAction
}

// NewDeferredRegistry creates a new DeferredRegistry.
func NewDeferredRegistry() *DeferredRegistry {
	return &DeferredRegistry{}
}

// Add registers a deferred action. Returns an error if the limit is exceeded.
func (r *DeferredRegistry) Add(action DeferredAction) error {
	r.mu.Lock()
	defer r.mu.Unlock()
	if len(r.actions) >= MaxDeferredActions {
		return fmt.Errorf("deferred action limit (%d) exceeded", MaxDeferredActions)
	}
	r.actions = append(r.actions, action)
	return nil
}

// Actions returns all registered deferred actions in registration order.
func (r *DeferredRegistry) Actions() []DeferredAction {
	r.mu.Lock()
	defer r.mu.Unlock()
	result := make([]DeferredAction, len(r.actions))
	copy(result, r.actions)
	return result
}

// Reset clears all registered deferred actions.
func (r *DeferredRegistry) Reset() {
	r.mu.Lock()
	r.actions = r.actions[:0]
	r.mu.Unlock()
}
