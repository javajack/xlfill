GO ?= go

.PHONY: test test-race bench lint cover clean

test:
	$(GO) test ./... -count=1 -timeout 60s

test-race:
	$(GO) test ./... -race -count=1 -timeout 120s

bench:
	$(GO) test ./... -bench=. -benchmem -run=^$$ -count=1

lint:
	$(GO) vet ./...

cover:
	$(GO) test ./... -coverprofile=cover.out -count=1
	$(GO) tool cover -func=cover.out | tail -1

clean:
	rm -f cover.out
