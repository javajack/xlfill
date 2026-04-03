package xlfill

import (
	"sync"
	"testing"

	"github.com/stretchr/testify/assert"
	"github.com/stretchr/testify/require"
)

func TestDeferredRegistry_AddAndActions(t *testing.T) {
	reg := NewDeferredRegistry()

	action1 := DeferredAction{
		Name:     "table",
		Sheet:    "Sheet1",
		StartRow: 0, StartCol: 0,
		EndRow: 10, EndCol: 5,
		Execute: func(tx *ExcelizeTransformer) error { return nil },
	}
	action2 := DeferredAction{
		Name:     "chart",
		Sheet:    "Sheet1",
		StartRow: 0, StartCol: 0,
		EndRow: 20, EndCol: 3,
		Execute: func(tx *ExcelizeTransformer) error { return nil },
	}

	reg.Add(action1)
	reg.Add(action2)

	actions := reg.Actions()
	require.Len(t, actions, 2)
	assert.Equal(t, "table", actions[0].Name)
	assert.Equal(t, "chart", actions[1].Name)
	assert.Equal(t, 10, actions[0].EndRow)
	assert.Equal(t, 20, actions[1].EndRow)
}

func TestDeferredRegistry_Reset(t *testing.T) {
	reg := NewDeferredRegistry()
	reg.Add(DeferredAction{Name: "test"})
	require.Len(t, reg.Actions(), 1)

	reg.Reset()
	assert.Empty(t, reg.Actions())
}

func TestDeferredRegistry_ActionsReturnsDefensiveCopy(t *testing.T) {
	reg := NewDeferredRegistry()
	reg.Add(DeferredAction{Name: "a"})

	actions := reg.Actions()
	actions[0].Name = "modified"

	// Original should be unmodified
	original := reg.Actions()
	assert.Equal(t, "a", original[0].Name)
}

func TestDeferredRegistry_ConcurrentAccess(t *testing.T) {
	reg := NewDeferredRegistry()
	var wg sync.WaitGroup

	for i := 0; i < 100; i++ {
		wg.Add(1)
		go func(n int) {
			defer wg.Done()
			reg.Add(DeferredAction{Name: "action", StartRow: n})
		}(i)
	}
	wg.Wait()

	actions := reg.Actions()
	assert.Len(t, actions, 100)
}

func TestContext_RegisterDeferred(t *testing.T) {
	ctx := NewContext(map[string]any{"x": 1})

	executed := false
	ctx.RegisterDeferred(DeferredAction{
		Name:  "test",
		Sheet: "Sheet1",
		Execute: func(tx *ExcelizeTransformer) error {
			executed = true
			return nil
		},
	})

	actions := ctx.Deferred().Actions()
	require.Len(t, actions, 1)
	assert.Equal(t, "test", actions[0].Name)

	// Execute the action
	err := actions[0].Execute(nil)
	require.NoError(t, err)
	assert.True(t, executed)
}

func TestContext_DeferredSharedAcrossClones(t *testing.T) {
	ctx := NewContext(map[string]any{})
	clone := ctx.Clone()

	ctx.RegisterDeferred(DeferredAction{Name: "from-original"})
	clone.RegisterDeferred(DeferredAction{Name: "from-clone"})

	// Both should see all actions (shared registry)
	assert.Len(t, ctx.Deferred().Actions(), 2)
	assert.Len(t, clone.Deferred().Actions(), 2)
}

func TestDeferredRegistry_EmptyActions(t *testing.T) {
	reg := NewDeferredRegistry()
	actions := reg.Actions()
	assert.NotNil(t, actions)
	assert.Empty(t, actions)
}

func TestDeferredAction_NilExecute(t *testing.T) {
	// Actions with nil Execute should not panic
	action := DeferredAction{Name: "noop", Execute: nil}
	assert.Nil(t, action.Execute)
}

func TestDeferredRegistry_OrderPreservation(t *testing.T) {
	reg := NewDeferredRegistry()
	for i := 0; i < 10; i++ {
		reg.Add(DeferredAction{Name: "action", StartRow: i})
	}

	actions := reg.Actions()
	for i, a := range actions {
		assert.Equal(t, i, a.StartRow)
	}
}
