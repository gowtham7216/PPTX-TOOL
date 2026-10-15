"""
SlideGuard Layout Region & Column Analyzer.
Analyzes spatial distribution of elements on a slide to identify header zones,
multi-column regions, side-by-side cards, and footer zones.
"""

from typing import List, Dict, Any, Tuple
from dataclasses import dataclass, field
from rules.reading_order.semantic_classifier import SemanticItem, SemanticRole


@dataclass
class ColumnRegion:
    """Represents a spatial column on a slide."""
    column_index: int  # 0-indexed left to right
    min_x: float
    max_x: float
    items: List[SemanticItem] = field(default_factory=list)

    @property
    def center_x(self) -> float:
        return (self.min_x + self.max_x) / 2.0


@dataclass
class SlideLayoutStructure:
    """Segmented layout zones for a slide."""
    header_items: List[SemanticItem] = field(default_factory=list)
    body_columns: List[ColumnRegion] = field(default_factory=list)
    footer_items: List[SemanticItem] = field(default_factory=list)
    unassigned_items: List[SemanticItem] = field(default_factory=list)
    is_multi_column: bool = False
    detected_layout_type: str = "SINGLE_COLUMN"  # "SINGLE_COLUMN", "TWO_COLUMN", "THREE_COLUMN", "MIXED"


class LayoutRegionAnalyzer:
    """
    Analyzes visual layout geometry to identify columns and structural regions.
    Prevents simplistic zig-zag ordering across multi-column slides.
    """

    def __init__(self, column_gap_threshold_pt: float = 24.0):
        self.column_gap_threshold_pt = column_gap_threshold_pt

    def analyze_layout(
        self,
        items: List[SemanticItem],
        slide_w_pt: float,
        slide_h_pt: float
    ) -> SlideLayoutStructure:
        layout = SlideLayoutStructure()

        if not items:
            return layout

        header_items = []
        footer_items = []
        body_candidates = []

        # 1. Separate Header (Titles), Footers, and Decorative Backgrounds
        for item in items:
            if item.role == SemanticRole.DECORATIVE:
                layout.unassigned_items.append(item)
            elif item.is_footer or item.role in (SemanticRole.FOOTER, SemanticRole.SLIDE_NUMBER):
                footer_items.append(item)
            elif item.is_title or (item.role == SemanticRole.TITLE and item.top < slide_h_pt * 0.25):
                header_items.append(item)
            else:
                body_candidates.append(item)


        layout.header_items = sorted(header_items, key=lambda x: (x.top, x.left))
        layout.footer_items = sorted(footer_items, key=lambda x: (x.top, x.left))

        if not body_candidates:
            return layout

        # 2. Check if body candidates form a horizontal timeline or alternating process flow
        timeline_candidates = [
            it for it in body_candidates
            if it.top < slide_h_pt * 0.85 and (it.top + it.height) < slide_h_pt * 0.90
        ]
        remaining_candidates = [it for it in body_candidates if it not in timeline_candidates]
        if self._is_horizontal_timeline_or_process(timeline_candidates, slide_w_pt, slide_h_pt):
            layout.detected_layout_type = "TIMELINE_HORIZONTAL"
            layout.is_multi_column = True
            # Sort strictly by horizontal progression (left coordinate) to preserve chronological timeline sequence
            timeline_items = sorted(timeline_candidates, key=lambda x: x.left)
            # Append any remaining bottom content items (e.g. disclaimer/footer text)
            sorted_remaining = sorted(remaining_candidates, key=lambda x: (x.top, x.left))
            all_ordered_items = timeline_items + sorted_remaining
            columns = [
                ColumnRegion(
                    column_index=idx,
                    min_x=it.left,
                    max_x=it.right,
                    items=[it]
                )
                for idx, it in enumerate(all_ordered_items)
            ]
            layout.body_columns = columns
            return layout

        # 3. Detect Columns among body candidates
        # Check horizontal distribution: compute X-overlap or X-clustering
        columns = self._cluster_into_columns(body_candidates, slide_w_pt)

        if len(columns) == 1:
            layout.detected_layout_type = "SINGLE_COLUMN"
            layout.is_multi_column = False
            # Sort single column strictly top-down
            columns[0].items.sort(key=lambda x: (x.top, x.left))
            layout.body_columns = columns
        elif len(columns) == 2:
            layout.detected_layout_type = "TWO_COLUMN"
            layout.is_multi_column = True
            for col in columns:
                col.items.sort(key=lambda x: (x.top, x.left))
            layout.body_columns = columns
        elif len(columns) == 3:
            layout.detected_layout_type = "THREE_COLUMN"
            layout.is_multi_column = True
            for col in columns:
                col.items.sort(key=lambda x: (x.top, x.left))
            layout.body_columns = columns
        else:
            layout.detected_layout_type = "MIXED"
            layout.is_multi_column = False
            # Fallback to single body column sorted top-down
            all_body = list(body_candidates)
            all_body.sort(key=lambda x: (x.top, x.left))
            col = ColumnRegion(column_index=0, min_x=0.0, max_x=slide_w_pt, items=all_body)
            layout.body_columns = [col]

        return layout

    def _is_horizontal_timeline_or_process(
        self,
        items: List[SemanticItem],
        slide_w_pt: float,
        slide_h_pt: float
    ) -> bool:
        """
        Detects whether items represent a horizontal timeline, workflow, or process cards.
        Key criteria:
        - At least 3 content cards with text
        - Progressively advancing horizontal centers left-to-right
        - No substantial vertical stacking between adjacent cards (horizontal overlap <= 40%)
        - Horizontal span across at least 40% of the slide width
        """
        if len(items) < 3:
            return False

        sorted_by_left = sorted(items, key=lambda x: x.left)
        centers_x = [it.left + it.width / 2.0 for it in sorted_by_left]

        # Centers must strictly advance left to right
        if not all(centers_x[i] < centers_x[i+1] for i in range(len(centers_x) - 1)):
            return False

        # Check for vertical stacking (columns): if any two adjacent cards overlap horizontally > 40%,
        # they are stacked vertically in a column, not in a horizontal timeline.
        for i in range(len(sorted_by_left) - 1):
            c1 = sorted_by_left[i]
            c2 = sorted_by_left[i + 1]
            overlap_x = max(0.0, min(c1.right, c2.right) - max(c1.left, c2.left))
            min_width = min(c1.width, c2.width)
            if min_width > 0 and (overlap_x / min_width) > 0.40:
                return False

        min_x = min(it.left for it in items)
        max_x = max(it.right for it in items)
        if (max_x - min_x) < slide_w_pt * 0.40:
            return False

        return True

    def _cluster_into_columns(
        self,
        items: List[SemanticItem],
        slide_w_pt: float
    ) -> List[ColumnRegion]:
        if len(items) <= 1:
            col = ColumnRegion(
                column_index=0,
                min_x=items[0].left if items else 0.0,
                max_x=items[0].right if items else slide_w_pt,
                items=list(items)
            )
            return [col]

        # Analyze horizontal midpoints
        # Check if items clearly divide into 2 sides (e.g. Left half < 48% width, Right half > 52% width)
        mid_x = slide_w_pt * 0.50
        left_items = []
        right_items = []
        spanning_items = []

        for item in items:
            # If item crosses over 60% of the slide width, it's a spanning element
            if item.width > slide_w_pt * 0.65:
                spanning_items.append(item)
            elif item.right <= mid_x + 36.0:  # clearly in left half with small margin
                left_items.append(item)
            elif item.left >= mid_x - 36.0:   # clearly in right half
                right_items.append(item)
            else:
                # Straddling center
                if item.left + (item.width / 2.0) < mid_x:
                    left_items.append(item)
                else:
                    right_items.append(item)

        # If both left and right columns have meaningful content and spanning items are few:
        if left_items and right_items:
            col_left = ColumnRegion(
                column_index=0,
                min_x=min(i.left for i in left_items),
                max_x=max(i.right for i in left_items),
                items=left_items
            )
            col_right = ColumnRegion(
                column_index=1,
                min_x=min(i.right for i in right_items),
                max_x=max(i.right for i in right_items),
                items=right_items
            )
            # If there are spanning items, return single column or handle appropriately
            if not spanning_items:
                return [col_left, col_right]

        # Single column fallback
        min_x = min(i.left for i in items)
        max_x = max(i.right for i in items)
        return [ColumnRegion(column_index=0, min_x=min_x, max_x=max_x, items=list(items))]
