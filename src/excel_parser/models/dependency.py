"""
Dependency graph DTOs for formula references.

Models the edges of the formula dependency graph: which cells reference
which other cells or ranges. Supports cross-sheet, external, structured,
and named range references. Used for upstream/downstream retrieval in RAG.
"""

from __future__ import annotations

from pydantic import Field

from .common import CellCoord, CellRange, EdgeType, StableModel, addr_to_a1, compute_hash


class DependencyEdgeDTO(StableModel):
    """
    A single dependency edge in the formula graph.

    Represents that `source_cell` (the cell with the formula) depends on
    `target_ref` (the cell or range it references). Edges are typed to
    distinguish local vs cross-sheet vs external references.
    """

    model_config = {"frozen": False, "extra": "forbid"}

    # Source: the cell containing the formula
    source_sheet: str
    source_coord: CellCoord

    # Target: what the formula references
    target_sheet: str | None = None  # None means same sheet as source
    target_coord: CellCoord | None = None  # For cell-to-cell refs
    target_range: CellRange | None = None  # For cell-to-range refs
    target_ref_string: str = ""  # Original reference string from formula

    # Classification
    edge_type: EdgeType = EdgeType.CELL_TO_CELL

    # For external references
    external_workbook: str | None = None

    # For named range references
    named_range_name: str | None = None

    # ID
    edge_id: str = Field(default="")

    def finalize(self) -> None:
        """Compute stable edge ID."""
        target = (
            self.target_ref_string
            or (self.target_coord.to_a1() if self.target_coord else "")
            or (self.target_range.to_a1() if self.target_range else "")
        )
        self.edge_id = compute_hash(
            self.source_sheet,
            self.source_coord.to_a1(),
            self.target_sheet or self.source_sheet,
            target,
        )

    @property
    def resolved_target_sheet(self) -> str:
        """Return the actual target sheet name."""
        return self.target_sheet or self.source_sheet


class DependencyGraph(StableModel):
    """
    Complete dependency graph for a workbook.

    Stores all edges and provides traversal methods for upstream/downstream
    lookups with cycle protection.
    """

    model_config = {"frozen": False, "extra": "forbid"}

    edges: list[DependencyEdgeDTO] = Field(default_factory=list)

    # Indexes built after all edges are added
    _forward: dict[str, list[DependencyEdgeDTO]] = {}  # source → [edges]
    _backward: dict[str, list[DependencyEdgeDTO]] = {}  # target → [edges]
    _circular_refs: set[str] = set()  # Cell IDs involved in circular references

    def model_post_init(self, __context) -> None:
        """Initialize internal indexes."""
        object.__setattr__(self, "_forward", {})
        object.__setattr__(self, "_backward", {})
        object.__setattr__(self, "_circular_refs", set())

    def add_edge(self, edge: DependencyEdgeDTO) -> None:
        """Add an edge and update indexes."""
        edge.finalize()
        self.edges.append(edge)
        source_key = f"{edge.source_sheet}!{edge.source_coord.to_a1()}"
        self._forward.setdefault(source_key, []).append(edge)

        if edge.target_coord:
            target_key = f"{edge.resolved_target_sheet}!{edge.target_coord.to_a1()}"
            self._backward.setdefault(target_key, []).append(edge)

    def build_indexes(self) -> None:
        """Rebuild forward and backward indexes from edge list."""
        self._forward.clear()
        self._backward.clear()
        for edge in self.edges:
            source_key = f"{edge.source_sheet}!{edge.source_coord.to_a1()}"
            self._forward.setdefault(source_key, []).append(edge)
            if edge.target_coord:
                target_key = f"{edge.resolved_target_sheet}!{edge.target_coord.to_a1()}"
                self._backward.setdefault(target_key, []).append(edge)

    def get_upstream(
        self, sheet: str, coord: CellCoord, max_depth: int = 5
    ) -> list[DependencyEdgeDTO]:
        """
        Get all upstream dependencies of a cell up to max_depth.

        Returns edges where the given cell is the source (i.e., cells
        that this cell's formula references). Includes transitive deps
        with cycle protection.
        """
        result = []
        visited: set[str] = set()
        self._traverse_upstream(sheet, coord, max_depth, 0, visited, result)
        return result

    def _traverse_upstream(
        self,
        sheet: str,
        coord: CellCoord,
        max_depth: int,
        current_depth: int,
        visited: set[str],
        result: list[DependencyEdgeDTO],
    ) -> None:
        """Recursive upstream traversal with cycle detection."""
        if current_depth >= max_depth:
            return
        key = f"{sheet}!{coord.to_a1()}"
        if key in visited:
            self._circular_refs.add(key)
            return
        visited.add(key)

        for edge in self._forward.get(key, []):
            result.append(edge)
            if edge.target_coord:
                self._traverse_upstream(
                    edge.resolved_target_sheet,
                    edge.target_coord,
                    max_depth,
                    current_depth + 1,
                    visited,
                    result,
                )

    def get_downstream(
        self, sheet: str, coord: CellCoord, max_depth: int = 5
    ) -> list[DependencyEdgeDTO]:
        """
        Get all downstream dependents of a cell up to max_depth.

        Returns edges where the given cell is a target (i.e., cells
        whose formulas reference this cell).
        """
        result = []
        visited: set[str] = set()
        self._traverse_downstream(sheet, coord, max_depth, 0, visited, result)
        return result

    def _traverse_downstream(
        self,
        sheet: str,
        coord: CellCoord,
        max_depth: int,
        current_depth: int,
        visited: set[str],
        result: list[DependencyEdgeDTO],
    ) -> None:
        """Recursive downstream traversal with cycle detection."""
        if current_depth >= max_depth:
            return
        key = f"{sheet}!{coord.to_a1()}"
        if key in visited:
            self._circular_refs.add(key)
            return
        visited.add(key)

        for edge in self._backward.get(key, []):
            result.append(edge)
            self._traverse_downstream(
                edge.source_sheet,
                edge.source_coord,
                max_depth,
                current_depth + 1,
                visited,
                result,
            )

    def detect_circular_refs(self) -> set[str]:
        """
        Detect all cells involved in circular references.

        Uses iterative Tarjan's SCC algorithm: single O(V+E) pass over the
        graph. A node is in a cycle iff it lives in a strongly-connected
        component of size >1, or a size-1 component with a self-loop.

        Returns set of cell keys (e.g., "Sheet1!A1") participating in cycles.
        """
        adj: dict[str, list[str]] = {}
        nodes: set[str] = set()
        for source_key, edges in self._forward.items():
            nodes.add(source_key)
            targets: list[str] = []
            for edge in edges:
                if edge.target_coord is None:
                    continue
                target_key = (
                    f"{edge.resolved_target_sheet}!"
                    f"{addr_to_a1(edge.target_coord.row, edge.target_coord.col)}"
                )
                targets.append(target_key)
                nodes.add(target_key)
            if targets:
                adj[source_key] = targets

        index_of: dict[str, int] = {}
        lowlink: dict[str, int] = {}
        on_stack: set[str] = set()
        stack: list[str] = []
        circular: set[str] = set()
        counter = [0]

        def strongconnect(root: str) -> None:
            # Iterative DFS. Python's default recursion limit would blow up
            # on deep formula chains (see adversarial test fixtures).
            work: list[tuple[str, int]] = [(root, 0)]
            index_of[root] = counter[0]
            lowlink[root] = counter[0]
            counter[0] += 1
            stack.append(root)
            on_stack.add(root)

            while work:
                node, next_i = work[-1]
                neighbors = adj.get(node, ())
                if next_i < len(neighbors):
                    work[-1] = (node, next_i + 1)
                    nb = neighbors[next_i]
                    if nb not in index_of:
                        index_of[nb] = counter[0]
                        lowlink[nb] = counter[0]
                        counter[0] += 1
                        stack.append(nb)
                        on_stack.add(nb)
                        work.append((nb, 0))
                    elif nb in on_stack:
                        if index_of[nb] < lowlink[node]:
                            lowlink[node] = index_of[nb]
                    continue

                work.pop()
                if work:
                    parent = work[-1][0]
                    if lowlink[node] < lowlink[parent]:
                        lowlink[parent] = lowlink[node]

                if lowlink[node] == index_of[node]:
                    scc: list[str] = []
                    while True:
                        w = stack.pop()
                        on_stack.discard(w)
                        scc.append(w)
                        if w == node:
                            break
                    if len(scc) > 1:
                        circular.update(scc)
                    else:
                        only = scc[0]
                        if only in adj and only in adj[only]:
                            circular.add(only)

        for v in nodes:
            if v not in index_of:
                strongconnect(v)

        object.__setattr__(self, "_circular_refs", circular)
        return circular

    @property
    def has_circular_refs(self) -> bool:
        return len(self._circular_refs) > 0
