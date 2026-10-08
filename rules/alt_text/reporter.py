"""
SlideGuard Task 3 Alternative Text & Decorative Image Reporting.
Constructs structured audit reports conforming strictly to Task 3 requirements:
  - Aggregate metric counts (meaningful, decorative, unknown, protected, etc.)
  - Per-image detailed verification status and audit lineage
"""

from typing import Dict, Any, List, Optional
from rules.alt_text.models import ImageDecision, ImageAction, ImageClassification


def generate_task_3_report(
    decisions: List[ImageDecision],
    actions_applied: Optional[List[Dict[str, Any]]] = None
) -> Dict[str, Any]:
    """
    Generates a structured Task 3 report dictionary containing:
      Total image objects, Meaningful images, Decorative images, Unknown images,
      Existing good alt text, Existing missing alt text, Generated alt text,
      Images marked decorative, Manual review items, High-confidence decisions,
      Medium-confidence decisions, Low-confidence decisions, Animation-protected images,
      Group-protected images, Failed images, and detailed per-image report.
    """
    total_images = len(decisions)
    meaningful = sum(1 for d in decisions if d.classification == ImageClassification.MEANINGFUL)
    decorative = sum(1 for d in decisions if d.classification == ImageClassification.DECORATIVE)
    unknown = sum(1 for d in decisions if d.classification == ImageClassification.UNKNOWN)

    existing_good = sum(1 for d in decisions if d.action == ImageAction.PRESERVE and bool(d.original_alt_text))
    existing_missing = sum(1 for d in decisions if not d.original_alt_text or d.action == ImageAction.GENERATE_ALT_TEXT)

    generated_alt = sum(1 for d in decisions if d.action == ImageAction.GENERATE_ALT_TEXT)
    marked_decorative = sum(1 for d in decisions if d.action == ImageAction.MARK_DECORATIVE)
    manual_review = sum(1 for d in decisions if d.action == ImageAction.MANUAL_REVIEW)

    high_conf = sum(1 for d in decisions if d.confidence >= 0.85)
    med_conf = sum(1 for d in decisions if 0.70 <= d.confidence < 0.85)
    low_conf = sum(1 for d in decisions if d.confidence < 0.70)

    anim_protected = sum(1 for d in decisions if d.is_animated)
    group_protected = sum(1 for d in decisions if d.is_grouped)
    failed = sum(1 for d in decisions if d.status == "FAILED")

    # Quality class counts
    specific_count = sum(1 for d in decisions if d.quality_class == "SPECIFIC")
    contextual_count = sum(1 for d in decisions if d.quality_class == "CONTEXTUAL")
    generic_count = sum(1 for d in decisions if d.quality_class == "GENERIC")
    manual_review_qc = sum(1 for d in decisions if d.quality_class == "MANUAL_REVIEW")

    # Generation basis counts
    visual_analysis_count = sum(1 for d in decisions if d.generation_basis == "VISUAL_ANALYSIS")
    structural_context_count = sum(1 for d in decisions if d.generation_basis == "STRUCTURAL_CONTEXT")
    existing_author_count = sum(1 for d in decisions if d.generation_basis == "EXISTING_AUTHOR_TEXT")
    deterministic_basis_count = sum(1 for d in decisions if d.generation_basis == "DETERMINISTIC")

    # Build per-image detail records
    per_image_details = []
    applied_map = {a.get("shape_id"): a for a in (actions_applied or [])}

    for d in decisions:
        applied_info = applied_map.get(d.shape_id, {})
        verif_status = "VERIFIED_APPLIED" if applied_info.get("success") else ("UNMODIFIED" if d.action == ImageAction.PRESERVE else ("PENDING_APPROVAL" if d.status == "PENDING_REVIEW" else d.status))

        per_image_details.append({
            "slide_number": d.slide_number,
            "shape_id": d.shape_id,
            "shape_name": d.shape_name,
            "classification": d.classification.value if hasattr(d.classification, "value") else str(d.classification),
            "action": d.action.value if hasattr(d.action, "value") else str(d.action),
            "original_alt_text": d.original_alt_text,
            "new_alt_text": d.effective_alt_text,
            "confidence": round(d.confidence, 2),
            "reason": d.reason,
            "generation_basis": d.generation_basis,
            "quality_class": d.quality_class,
            "verification_status": verif_status,
            "is_animated": d.is_animated,
            "is_grouped": d.is_grouped,
            "model_traceability": d.model_traceability,
        })

    return {
        "summary": {
            "total_images": total_images,
            "meaningful_images": meaningful,
            "decorative_images": decorative,
            "unknown_images": unknown,
            "existing_good_alt_text": existing_good,
            "existing_missing_alt_text": existing_missing,
            "generated_alt_text": generated_alt,
            "images_marked_decorative": marked_decorative,
            "manual_review_items": manual_review,
            "high_confidence_decisions": high_conf,
            "medium_confidence_decisions": med_conf,
            "low_confidence_decisions": low_conf,
            "animation_protected_images": anim_protected,
            "group_protected_images": group_protected,
            "failed_images": failed,
            "quality_specific": specific_count,
            "quality_contextual": contextual_count,
            "quality_generic": generic_count,
            "quality_manual_review": manual_review_qc,
            "basis_visual_analysis": visual_analysis_count,
            "basis_structural_context": structural_context_count,
            "basis_existing_author": existing_author_count,
            "basis_deterministic": deterministic_basis_count,
        },
        "changed_images": per_image_details,
    }
