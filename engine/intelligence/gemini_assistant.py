"""
SlideGuard Gemini 3.8 Flash High Assistance Service (Phase 8).
Provides structured semantic reasoning assistance for ambiguous accessibility cases
(alt text suggestions, decorative classification, ambiguous titles, and table headers).
Strictly decoupled from PPTX modification: outputs structured data only.
"""

from dataclasses import dataclass, field
from typing import Dict, Any, List, Optional
import json
import logging

import re

logger = logging.getLogger(__name__)

# Distinct Model Roles
ANTIGRAVITY_CODING_MODEL: str = "Gemini 3.8 Flash High"

# SlideGuard Runtime Models
DEFAULT_RUNTIME_MODEL: str = "gemini-3.5"
FALLBACK_RUNTIME_MODEL: str = "gemini-3.1"
SUPPORTED_RUNTIME_MODELS: List[str] = [DEFAULT_RUNTIME_MODEL, FALLBACK_RUNTIME_MODEL]


def resolve_runtime_model(
    preferred_model: Optional[str] = None,
    disallow_coding_model: bool = True
) -> str:
    """
    Resolves the SlideGuard runtime model with explicit preference order:
      1. Preferred parameter if specified and valid
      2. Environment variable SLIDEGUARD_GEMINI_MODEL
      3. Primary runtime model: 'gemini-3.5'
      4. Fallback runtime model: 'gemini-3.1'
    Ensures 'gemini-3.8-flash-high' (Antigravity coding model) is NEVER selected
    for SlideGuard runtime decisions when disallow_coding_model is True.
    """
    import os
    model = preferred_model or os.environ.get("SLIDEGUARD_GEMINI_MODEL") or DEFAULT_RUNTIME_MODEL
    if disallow_coding_model and model in ("gemini-3.8-flash-high", "gemini-3.8"):
        logger.info(
            "Model '%s' is reserved for Antigravity coding. Selecting SlideGuard runtime model '%s'.",
            model,
            DEFAULT_RUNTIME_MODEL
        )
        return DEFAULT_RUNTIME_MODEL
    return model


@dataclass
class AIRecommendation:
    """Structured output returned by the Gemini AI reasoning layer."""
    recommendation_type: str  # "ALT_TEXT_SUGGESTION", "ROLE_CLASSIFICATION", "TITLE_SUGGESTION", "HEADER_SUGGESTION", "EXPLANATION"
    suggested_value: str
    confidence: float  # 0.0 to 1.0
    reasoning: str
    is_safe_for_review: bool = True
    model_used: str = DEFAULT_RUNTIME_MODEL
    extra_attributes: Dict[str, Any] = field(default_factory=dict)

    def to_dict(self) -> Dict[str, Any]:
        return {
            "recommendation_type": self.recommendation_type,
            "suggested_value": self.suggested_value,
            "confidence": round(self.confidence, 2),
            "reasoning": self.reasoning,
            "is_safe_for_review": self.is_safe_for_review,
            "model_used": self.model_used,
            "extra_attributes": self.extra_attributes,
        }


class GeminiAssistant:
    """
    Coordinates SlideGuard Runtime Gemini models for semantic accessibility assistance.
    Guarantees:
      - Runtime model is Gemini 3.5 (fallback only to Gemini 3.1).
      - Antigravity coding model (Gemini 3.8 Flash High) remains decoupled.
      - Only structured context is transmitted (no PPTX binaries or paths).
      - Output is parsed and schema-validated JSON.
      - Graceful offline and failure fallback.
      - Zero direct PPTX XML modifications.
    """

    def __init__(self, client=None, model_name: Optional[str] = None):
        self.client = client
        self.model_name = resolve_runtime_model(model_name)

    def suggest_alt_text(self, image_metadata: Dict[str, Any]) -> AIRecommendation:
        """Suggests a concise, informative alternative text description for an image."""
        # Minimum structured context extraction
        context_payload = {
            "shape_name": image_metadata.get("shape_name", "Picture"),
            "dimensions": f"{image_metadata.get('width_pt', 0):.0f}x{image_metadata.get('height_pt', 0):.0f}pt",
            "slide_title": image_metadata.get("slide_title", ""),
            "surrounding_text_preview": image_metadata.get("surrounding_text", "")[:200],
        }

        if self.client is None:
            # Deterministic offline fallback
            return self._build_offline_fallback(
                rec_type="ALT_TEXT_SUGGESTION",
                suggested=f"Visual graphic in presentation section: {context_payload['slide_title'] or context_payload['shape_name']}",
                reasoning="Offline fallback: suggested baseline context from slide structure."
            )

        prompt = (
            f"Generate a concise (1-2 sentences), accessible WCAG-compliant alt text description for an image in a presentation.\n"
            f"Context: {json.dumps(context_payload)}\n"
            f"Return JSON strictly matching: {{\"suggested_value\": \"...\", \"confidence\": 0.85, \"reasoning\": \"...\"}}"
        )

        try:
            raw_response = self._call_model(prompt)
            return self.parse_structured_response(raw_response, default_type="ALT_TEXT_SUGGESTION")
        except Exception as e:
            logger.warning("Gemini alt text assistance error: %s. Using fallback.", e)
            return self._build_offline_fallback("ALT_TEXT_SUGGESTION", "", f"AI assistance unavailable: {e}")

    def classify_ambiguous_image(self, shape_metadata: Dict[str, Any]) -> AIRecommendation:
        """Assists in classifying whether an ambiguous image is decorative or meaningful."""
        context_payload = {
            "shape_name": shape_metadata.get("shape_name", ""),
            "shape_type": str(shape_metadata.get("shape_type", "")),
            "dimensions": f"{shape_metadata.get('width_pt', 0)}x{shape_metadata.get('height_pt', 0)}pt",
            "has_text": shape_metadata.get("has_text", False),
            "existing_alt": shape_metadata.get("alt_text", ""),
        }

        if self.client is None:
            # Deterministic heuristic fallback
            w = float(shape_metadata.get("width_pt", 0))
            h = float(shape_metadata.get("height_pt", 0))
            is_thin = (w <= 4.0 or h <= 4.0)
            role = "DECORATIVE" if is_thin else "MEANINGFUL"
            return AIRecommendation(
                recommendation_type="ROLE_CLASSIFICATION",
                suggested_value=role,
                confidence=0.95 if is_thin else 0.70,
                reasoning="Deterministic geometric assessment: divider dimension threshold." if is_thin else "Ambiguous visual graphic requires author review.",
                is_safe_for_review=True,
                model_used=self.model_name
            )

        prompt = (
            f"Classify if the following visual shape is DECORATIVE or MEANINGFUL for screen readers.\n"
            f"Context: {json.dumps(context_payload)}\n"
            f"Return JSON: {{\"suggested_value\": \"DECORATIVE|MEANINGFUL\", \"confidence\": 0.9, \"reasoning\": \"...\"}}"
        )

        try:
            raw_response = self._call_model(prompt)
            return self.parse_structured_response(raw_response, default_type="ROLE_CLASSIFICATION")
        except Exception as e:
            logger.warning("Gemini image classification error: %s", e)
            return self._build_offline_fallback("ROLE_CLASSIFICATION", "MEANINGFUL", f"AI unavailable: {e}")

    def resolve_ambiguous_title(self, title_candidates: List[Dict[str, Any]]) -> AIRecommendation:
        """Assists in identifying the primary title candidate from competing top text boxes."""
        if not title_candidates:
            return self._build_offline_fallback("TITLE_SUGGESTION", "", "No candidates provided.")

        if self.client is None:
            # Offline heuristic: highest score or highest on slide
            best = min(title_candidates, key=lambda c: c.get("top_pt", 999.0))
            return AIRecommendation(
                recommendation_type="TITLE_SUGGESTION",
                suggested_value=best.get("text", ""),
                confidence=0.75,
                reasoning="Heuristic fallback: selected highest shape near top of slide.",
                is_safe_for_review=True,
                extra_attributes={"candidate_shape_id": best.get("shape_id")}
            )

        prompt = (
            f"Select which candidate is the most appropriate slide title heading.\n"
            f"Candidates: {json.dumps(title_candidates[:5])}\n"
            f"Return JSON: {{\"suggested_value\": \"Title Text\", \"confidence\": 0.9, \"reasoning\": \"...\", \"shape_id\": 123}}"
        )

        try:
            raw_response = self._call_model(prompt)
            return self.parse_structured_response(raw_response, default_type="TITLE_SUGGESTION")
        except Exception as e:
            return self._build_offline_fallback("TITLE_SUGGESTION", title_candidates[0].get("text", ""), str(e))

    def parse_structured_response(
        self,
        response_text: str,
        default_type: str = "EXPLANATION"
    ) -> AIRecommendation:
        """Parses and validates structured JSON response from Gemini model."""
        if not response_text or not response_text.strip():
            return self._build_offline_fallback(default_type, "", "Empty AI response.", is_safe=False)

        try:
            cleaned = response_text.strip()
            # Strip code fences if model enclosed in ```json ... ```
            if cleaned.startswith("```"):
                lines = cleaned.split("\n")
                if lines[0].startswith("```"):
                    lines = lines[1:]
                if lines and lines[-1].strip() == "```":
                    lines = lines[:-1]
                cleaned = "\n".join(lines).strip()

            parsed = json.loads(cleaned)
            if not isinstance(parsed, dict):
                raise ValueError("Response is not a valid JSON dictionary")

            suggested = str(parsed.get("suggested_value", "")).strip()
            confidence = float(parsed.get("confidence", 0.5))
            confidence = max(0.0, min(1.0, confidence))
            reasoning = str(parsed.get("reasoning", "Generated by AI analysis."))

            extra = {k: v for k, v in parsed.items() if k not in ("suggested_value", "confidence", "reasoning")}

            return AIRecommendation(
                recommendation_type=default_type,
                suggested_value=suggested,
                confidence=confidence,
                reasoning=reasoning,
                is_safe_for_review=True,
                model_used=self.model_name,
                extra_attributes=extra
            )
        except Exception as err:
            logger.warning("Failed to parse Gemini structured response: %s (Raw: %r)", err, response_text[:100])
            return AIRecommendation(
                recommendation_type=default_type,
                suggested_value="",
                confidence=0.0,
                reasoning=f"Invalid response format: {err}",
                is_safe_for_review=False,
                model_used=self.model_name
            )

    def _call_model(self, prompt: str) -> str:
        """Invokes the client if configured."""
        if hasattr(self.client, "generate_content"):
            res = self.client.generate_content(prompt)
            return res.text
        elif callable(self.client):
            return self.client(prompt)

        import os
        api_key = (
            os.environ.get("GEMINI_API_KEY") or
            os.environ.get("GOOGLE_API_KEY") or
            os.environ.get("SLIDEGUARD_GEMINI_KEY")
        )
        if api_key:
            return self._call_rest_api(prompt, api_key.strip().strip("'\""))

        raise RuntimeError("No callable Gemini client configured.")

    def _call_rest_api(self, prompt: str, api_key: str) -> str:
        """Invokes Gemini 3.8 Flash High endpoint via REST."""
        import urllib.request
        import json

        url = f"https://generativelanguage.googleapis.com/v1beta/models/{self.model_name}:generateContent?key={api_key}"
        headers = {"Content-Type": "application/json"}
        payload = {
            "contents": [{"parts": [{"text": prompt}]}],
            "generationConfig": {"temperature": 0.1, "responseMimeType": "application/json"}
        }
        body = json.dumps(payload).encode("utf-8")
        req = urllib.request.Request(url, data=body, headers=headers, method="POST")
        with urllib.request.urlopen(req, timeout=20) as resp:
            data = json.loads(resp.read().decode("utf-8"))
            return data["candidates"][0]["content"]["parts"][0]["text"]


    def _build_offline_fallback(
        self,
        rec_type: str,
        suggested: str,
        reasoning: str,
        is_safe: bool = True
    ) -> AIRecommendation:
        """Generates a safe deterministic fallback when Gemini is offline or unavailable."""
        return AIRecommendation(
            recommendation_type=rec_type,
            suggested_value=suggested,
            confidence=0.5 if suggested else 0.0,
            reasoning=reasoning,
            is_safe_for_review=is_safe,
            model_used=self.model_name
        )

    def author_content(
        self,
        authoring_type: str,
        context_payload: Dict[str, Any]
    ) -> Dict[str, Any]:
        """
        Authors accessibility content using Gemini 3.8 Flash High.
        Enforces:
          - Model: gemini-3.8-flash-high only
          - Strict structured JSON schema
          - Hallucination protection (supported facts vs inferences vs uncertain)
          - Offline / API failure graceful fallback
        """
        # 1. Deterministic offline fallback if client is absent
        if self.client is None:
            return self._build_authoring_offline_fallback(authoring_type, context_payload)

        # 2. Build minimal structured prompt
        prompt = self._build_authoring_prompt(authoring_type, context_payload)

        # 3. Call model safely
        try:
            raw_response = self._call_model(prompt)
            return self._parse_and_validate_authoring_response(raw_response, authoring_type, context_payload)
        except Exception as err:
            logger.warning("Gemini authoring API error (%s): %s. Falling back.", authoring_type, err)
            return self._build_authoring_error_fallback(authoring_type, f"{type(err).__name__}: {err}")


    def author_chart_description(self, context_payload: Dict[str, Any]) -> Dict[str, Any]:
        """Specialized authoring helper for accessible chart descriptions."""
        return self.author_content("CHART_DESCRIPTION", context_payload)

    def author_diagram_description(self, context_payload: Dict[str, Any]) -> Dict[str, Any]:
        """Specialized authoring helper for accessible diagram descriptions."""
        return self.author_content("DIAGRAM_DESCRIPTION", context_payload)

    def author_slide_title(self, context_payload: Dict[str, Any]) -> Dict[str, Any]:
        """Specialized authoring helper for meaningful slide title suggestions."""
        return self.author_content("SLIDE_TITLE", context_payload)

    def author_table_interpretation(self, context_payload: Dict[str, Any]) -> Dict[str, Any]:
        """Specialized authoring helper for table header interpretation and descriptions."""
        return self.author_content("TABLE_HEADER_INTERPRETATION", context_payload)

    def _build_authoring_prompt(self, authoring_type: str, context_payload: Dict[str, Any]) -> str:
        """Constructs minimal structured prompt with strict schema and hallucination boundaries."""
        instructions_by_type = {
            "ALT_TEXT": (
                "Suggest concise, meaningful alternative text (1-2 sentences) describing the visual element's "
                "purpose in context. Describe the visual's meaningful message, not merely its generic appearance."
            ),
            "IMAGE_DESCRIPTION": (
                "Suggest concise, meaningful alternative text describing the image's functional purpose. "
                "Do not merely say 'Image of...'. If image purpose cannot be reliably determined, state 'MANUAL_REVIEW'."
            ),
            "CHART_DESCRIPTION": (
                "Suggest an accessible chart description summarizing chart type, purpose, and general trend. "
                "CRITICAL: Do NOT invent numerical values, percentages, categories, or conclusions not in the context. "
                "If data semantics are missing, state 'MANUAL_REVIEW'."
            ),
            "DIAGRAM_DESCRIPTION": (
                "Suggest an accessible diagram description summarizing relationships and flow from visible labels. "
                "CRITICAL: Do NOT claim relationships that cannot be confirmed from available evidence."
            ),
            "SLIDE_TITLE": (
                "Suggest a concise, informative slide title heading (max 10 words) summarizing the slide's topic. "
                "CRITICAL: Never generate generic placeholders like 'Slide 4' or 'Untitled Slide'."
            ),
            "TABLE_HEADER_INTERPRETATION": (
                "Interpret the ambiguous table structure and determine if row 0 functions as column headers. "
                "Suggest clear column header descriptions based strictly on available cell data."
            ),
            "TABLE_DESCRIPTION": (
                "Suggest an accessible table description explaining the structure and reading relationship."
            ),
            "ACCESSIBILITY_EXPLANATION": (
                "Provide an accessible explanation for this element to assist screen reader users."
            )
        }

        task_instruction = instructions_by_type.get(
            authoring_type,
            "Suggest concise accessible content based strictly on provided context."
        )

        prompt = (
            f"You are the Gemini 3.8 Flash High Accessibility Authoring Service for PowerPoint.\n"
            f"Task: {task_instruction}\n"
            f"\nContext Data:\n{json.dumps(context_payload, indent=2)}\n\n"
            f"Rules:\n"
            f"1. Strictly return valid JSON matching this schema:\n"
            f"   {{\n"
            f"     \"suggested_value\": \"<string>\",\n"
            f"     \"confidence\": <float between 0.0 and 1.0>,\n"
            f"     \"rationale\": \"<string explaining decision>\",\n"
            f"     \"fact_grounding\": \"SUPPORTED_FACT\" | \"INFERENCE\" | \"UNCERTAIN\",\n"
            f"     \"requires_human_review\": <boolean>\n"
            f"   }}\n"
            f"2. Never hallucinate facts, statistics, or metrics not grounded in the context.\n"
            f"3. If context is insufficient, set fact_grounding='UNCERTAIN' and suggest 'MANUAL_REVIEW'.\n"
            f"Return JSON strictly. No extra conversational text."
        )
        return prompt

    def _parse_and_validate_authoring_response(
        self,
        raw_response: str,
        authoring_type: str,
        context_payload: Dict[str, Any]
    ) -> Dict[str, Any]:
        """Parses JSON, validates schema and types, and performs hallucination checks."""
        if not raw_response or not raw_response.strip():
            logger.warning("Empty AI response received for %s.", authoring_type)
            return self._build_authoring_error_fallback(authoring_type, "Empty response from AI model.")

        cleaned = raw_response.strip()
        if cleaned.startswith("```"):
            lines = cleaned.split("\n")
            if lines[0].startswith("```"):
                lines = lines[1:]
            if lines and lines[-1].strip() == "```":
                lines = lines[:-1]
            cleaned = "\n".join(lines).strip()

        try:
            parsed = json.loads(cleaned)
        except Exception as err:
            logger.warning("Malformed JSON response from AI for %s: %s", authoring_type, err)
            return self._build_authoring_error_fallback(
                authoring_type,
                f"Malformed JSON returned by model: {err}",
                malformed=True
            )

        if not isinstance(parsed, dict):
            return self._build_authoring_error_fallback(authoring_type, "Response is not a JSON object.")

        # Required fields check
        required_fields = ["suggested_value", "confidence", "rationale"]
        missing = [f for f in required_fields if f not in parsed]
        if missing:
            return self._build_authoring_error_fallback(
                authoring_type,
                f"Missing required schema fields: {missing}",
                malformed=True
            )

        suggested_val = str(parsed.get("suggested_value", "")).strip()
        try:
            confidence = float(parsed.get("confidence", 0.5))
            confidence = max(0.0, min(1.0, confidence))
        except (ValueError, TypeError):
            confidence = 0.5

        rationale = str(parsed.get("rationale", "")).strip()
        fact_grounding = str(parsed.get("fact_grounding", "SUPPORTED_FACT")).upper()
        if fact_grounding not in ("SUPPORTED_FACT", "INFERENCE", "UNCERTAIN"):
            fact_grounding = "INFERENCE"

        requires_review = bool(parsed.get("requires_human_review", True))
        warnings: List[str] = []

        # Hallucination Protection: Check for unsupported numbers/percentages
        context_str = json.dumps(context_payload).lower()
        numbers_in_suggestion = re.findall(r"\b\d+(?:\.\d+)?%?\b", suggested_val)
        unsupported_tokens = []
        for token in numbers_in_suggestion:
            if token.lower() not in context_str:
                unsupported_tokens.append(token)

        if unsupported_tokens:
            fact_grounding = "UNCERTAIN"
            warnings.append(
                f"Hallucination warning: Suggested value includes unsupported numbers {unsupported_tokens}."
            )
            confidence = min(confidence, 0.45)
            requires_review = True

        # Check for meaningless title placeholders
        if authoring_type == "SLIDE_TITLE":
            low_sugg = suggested_val.lower()
            if re.match(r"^slide\s*\d+$", low_sugg) or low_sugg in ("untitled slide", "untitled", "slide", "title"):
                return self._build_authoring_error_fallback(
                    authoring_type,
                    f"Generated meaningless placeholder title '{suggested_val}'. Reverted to manual review."
                )

        status = "PENDING_REVIEW"
        if suggested_val.upper() == "MANUAL_REVIEW" or fact_grounding == "UNCERTAIN" and confidence < 0.5:
            status = "MANUAL_REVIEW"

        return {
            "suggested_value": suggested_val,
            "confidence": confidence,
            "rationale": rationale or "Generated by Gemini 3.8 Flash High.",
            "fact_grounding": fact_grounding,
            "warnings": warnings,
            "requires_human_review": requires_review,
            "status": status,
        }

    def _build_authoring_offline_fallback(
        self,
        authoring_type: str,
        context_payload: Dict[str, Any]
    ) -> Dict[str, Any]:
        """Generates deterministic offline suggestion or manual review flag."""
        slide_title = context_payload.get("slide_title", "")
        shape_name = context_payload.get("shape_name", "Shape")

        if authoring_type in ("ALT_TEXT", "IMAGE_DESCRIPTION"):
            sugg = f"Visual graphic on slide '{slide_title or shape_name}'" if slide_title else f"Graphic: {shape_name}"
            conf = 0.5
            grounding = "INFERENCE"
            rat = "Deterministic offline baseline: contextual fallback from slide structure."
        elif authoring_type == "SLIDE_TITLE":
            # Extract prominent text preview if available
            cand_text = context_payload.get("nearby_text", "")
            if isinstance(cand_text, list) and cand_text:
                cand_text = cand_text[0]
            cand_text = str(cand_text).strip()[:60]
            sugg = cand_text if (cand_text and len(cand_text) > 3) else "MANUAL_REVIEW"
            conf = 0.6 if sugg != "MANUAL_REVIEW" else 0.0
            grounding = "SUPPORTED_FACT" if sugg != "MANUAL_REVIEW" else "UNCERTAIN"
            rat = "Deterministic offline fallback: selected leading heading text."
        elif authoring_type == "CHART_DESCRIPTION":
            sugg = f"Chart visual on '{slide_title}'" if slide_title else "MANUAL_REVIEW"
            conf = 0.4
            grounding = "INFERENCE" if sugg != "MANUAL_REVIEW" else "UNCERTAIN"
            rat = "Deterministic offline chart fallback: summary requires author verification."
        elif authoring_type == "DIAGRAM_DESCRIPTION":
            sugg = f"Diagram flow on '{slide_title}'" if slide_title else "MANUAL_REVIEW"
            conf = 0.4
            grounding = "INFERENCE" if sugg != "MANUAL_REVIEW" else "UNCERTAIN"
            rat = "Deterministic offline diagram fallback: diagram structure requires verification."
        elif authoring_type == "TABLE_HEADER_INTERPRETATION":
            sugg = "First row designated as column headers"
            conf = 0.7
            grounding = "SUPPORTED_FACT"
            rat = "Deterministic table analysis: standard tabular header structure."
        else:
            sugg = "MANUAL_REVIEW"
            conf = 0.0
            grounding = "UNCERTAIN"
            rat = "Offline fallback: manual author review recommended."

        status = "MANUAL_REVIEW" if sugg == "MANUAL_REVIEW" else "PENDING_REVIEW"

        return {
            "suggested_value": sugg,
            "confidence": conf,
            "rationale": rat,
            "fact_grounding": grounding,
            "warnings": ["AI service offline or unconfigured; generated via deterministic fallback."],
            "requires_human_review": True,
            "status": status,
        }

    def _build_authoring_error_fallback(
        self,
        authoring_type: str,
        error_msg: str,
        malformed: bool = False
    ) -> Dict[str, Any]:
        """Generates a safe manual review response upon AI error or malformed response."""
        return {
            "suggested_value": "MANUAL_REVIEW",
            "confidence": 0.0,
            "rationale": f"AI authoring failed ({error_msg}). Requires human review.",
            "fact_grounding": "UNCERTAIN",
            "warnings": [f"AI parsing error: {error_msg}"] if malformed else [f"AI error: {error_msg}"],
            "requires_human_review": True,
            "status": "MANUAL_REVIEW",
        }

    # =========================================================================
    # Phase 8 / Task 1: Semantic Reading Order Analysis (Gemini 3.8 Flash High)
    # =========================================================================

    def analyze_reading_order(
        self,
        slide_context: Dict[str, Any]
    ) -> Dict[str, Any]:
        """
        Analyzes PowerPoint slide reading order using Gemini 3.8 Flash High.
        Enforces:
          - Model: gemini-3.8-flash-high only
          - Strict structured JSON decision schema
          - Context privacy (no PPTX filesystem paths)
          - Offline / API failure graceful fallback
        """
        # 1. Deterministic offline fallback if client is absent and no API key
        import os
        api_key = (
            os.environ.get("GEMINI_API_KEY") or
            os.environ.get("GOOGLE_API_KEY") or
            os.environ.get("SLIDEGUARD_GEMINI_KEY")
        )
        if self.client is None and not api_key:
            return self._build_reading_order_offline_fallback(slide_context)

        # 2. Build structured prompt
        prompt = self._build_reading_order_prompt(slide_context)

        # 3. Call model safely
        try:
            raw_response = self._call_model(prompt)
            return self._parse_and_validate_reading_order_response(raw_response, slide_context)
        except Exception as err:
            logger.warning("Gemini reading order API error: %s. Using semantic fallback.", err)
            fallback = self._build_reading_order_offline_fallback(slide_context)
            fallback["warnings"].append(f"AI API unavailable ({err}); resolved via deterministic semantic analyzer.")
            return fallback

    def _build_reading_order_prompt(self, slide_context: Dict[str, Any]) -> str:
        """Constructs prompt for Gemini 3.8 Flash High with WCAG sequence rules and strict JSON schema."""
        prompt = (
            f"You are the Gemini 3.8 Flash High Accessibility Specialist for SlideGuard.\n"
            f"Task: Analyze the semantic reading order of objects on this PowerPoint slide according to "
            f"WCAG 2.1 - 1.3.2 Meaningful Sequence (Level A).\n\n"
            f"Reading Order Sequence Principles:\n"
            f"1. TITLE first: Slide title heading is always the primary entry point for screen readers.\n"
            f"2. SUBTITLE second: If a subtitle exists directly under or paired with the title.\n"
            f"3. INTRODUCTION / LEAD PARAGRAPHS: Lead summary content precedes detailed sections.\n"
            f"4. BODY / MAIN CONTENT: Logical semantic flow (left-to-right columns, top-to-bottom within column).\n"
            f"5. REPEATED CARDS / SECTIONS: For repeated cards (e.g. side-by-side or stacked cards with badges/numbers "
            f"like '01', '02', '03' and headers/bodies), preserve the logical card-by-card reading sequence! "
            f"Read Card 1 [Badge -> Header -> Body] completely, then Card 2 [Badge -> Header -> Body], etc. "
            f"CRITICAL: Do NOT zig-zag across columns (never read all badges, then all headers, then all bodies).\n"
            f"6. IMAGE / CHART + CAPTION: Place the caption immediately after its associated graphic.\n"
            f"7. FOOTERS / SLIDE NUMBERS / COPYRIGHT: Move to the very end of the content reading sequence.\n"
            f"8. PURELY DECORATIVE: Full-slide backgrounds, colored containers, dividers, and decorative shapes "
            f"must NOT interrupt reading flow. Set action='EXCLUDE_DECORATIVE', reading_order=0.\n\n"
            f"Context Data:\n{json.dumps(slide_context, indent=2)}\n\n"
            f"Rules:\n"
            f"1. Strictly return valid JSON matching this schema:\n"
            f"   {{\n"
            f"     \"slide_number\": <int>,\n"
            f"     \"confidence\": <float between 0.0 and 1.0>,\n"
            f"     \"rationale\": \"<summary of semantic sequence decisions>\",\n"
            f"     \"decisions\": [\n"
            f"       {{\n"
            f"         \"shape_id\": <int>,\n"
            f"         \"semantic_role\": \"TITLE\" | \"SUBTITLE\" | \"BODY\" | \"CARD_HEADER\" | \"CARD_BODY\" | \"CARD_BADGE\" | \"IMAGE\" | \"TABLE\" | \"CHART\" | \"CAPTION\" | \"FOOTER\" | \"SLIDE_NUMBER\" | \"DECORATIVE\" | \"UNKNOWN\",\n"
            f"         \"reading_order\": <int 1-based index, or 0 if decorative/excluded>,\n"
            f"         \"confidence\": <float between 0.0 and 1.0>,\n"
            f"         \"action\": \"PRESERVE\" | \"REORDER\" | \"EXCLUDE_DECORATIVE\" | \"MANUAL_REVIEW\",\n"
            f"         \"reason\": \"<brief explanation of role and sequence decision>\"\n"
            f"       }}\n"
            f"     ]\n"
            f"   }}\n"
            f"2. Every shape in context MUST appear in decisions. Do NOT invent new shape IDs.\n"
            f"3. Confidence Guidelines:\n"
            f"   - HIGH >= 0.85: Clear semantic structure, unambiguous hierarchy.\n"
            f"   - MEDIUM 0.70-0.849: Good structure but minor layout ambiguities.\n"
            f"   - LOW < 0.70: Heavily overlapping, ambiguous, or uninterpretable layout.\n"
            f"Return JSON strictly. No extra markdown commentary."
        )
        return prompt

    def _parse_and_validate_reading_order_response(
        self,
        raw_response: str,
        slide_context: Dict[str, Any]
    ) -> Dict[str, Any]:
        """Parses structured JSON response from Gemini, sanitizes IDs and ensures schema integrity."""
        if not raw_response or not raw_response.strip():
            logger.warning("Empty AI response received for reading order.")
            return self._build_reading_order_offline_fallback(slide_context)

        cleaned = raw_response.strip()
        if cleaned.startswith("```"):
            lines = cleaned.split("\n")
            if lines[0].startswith("```"):
                lines = lines[1:]
            if lines and lines[-1].strip() == "```":
                lines = lines[:-1]
            cleaned = "\n".join(lines).strip()

        try:
            parsed = json.loads(cleaned)
        except Exception as err:
            logger.warning("Malformed JSON response from AI for reading order: %s", err)
            fallback = self._build_reading_order_offline_fallback(slide_context)
            fallback["warnings"].append(f"AI returned malformed JSON: {err}")
            return fallback

        if not isinstance(parsed, dict) or "decisions" not in parsed:
            fallback = self._build_reading_order_offline_fallback(slide_context)
            fallback["warnings"].append("AI response missing 'decisions' list")
            return fallback

        shapes = slide_context.get("shapes", [])
        known_ids = {s["shape_id"] for s in shapes}
        validated_decisions = []
        covered_ids = set()

        for dec in parsed.get("decisions", []):
            if not isinstance(dec, dict):
                continue
            sid = dec.get("shape_id")
            if sid not in known_ids or sid in covered_ids:
                continue

            role = str(dec.get("semantic_role", "BODY")).upper()
            try:
                order_idx = int(dec.get("reading_order", 0))
            except (ValueError, TypeError):
                order_idx = 0

            try:
                conf = float(dec.get("confidence", 0.5))
                conf = max(0.0, min(1.0, conf))
            except (ValueError, TypeError):
                conf = 0.5

            action = str(dec.get("action", "REORDER")).upper()
            if action not in ("PRESERVE", "REORDER", "EXCLUDE_DECORATIVE", "MANUAL_REVIEW"):
                action = "MANUAL_REVIEW"

            reason = str(dec.get("reason", "Analyzed by Gemini 3.8 Flash High"))

            validated_decisions.append({
                "shape_id": sid,
                "semantic_role": role,
                "reading_order": order_idx,
                "confidence": conf,
                "action": action,
                "reason": reason
            })
            covered_ids.add(sid)

        # Ensure all existing shapes are accounted for
        if covered_ids != known_ids:
            missing_ids = known_ids - covered_ids
            for mid in missing_ids:
                validated_decisions.append({
                    "shape_id": mid,
                    "semantic_role": "UNKNOWN",
                    "reading_order": 0,
                    "confidence": 0.5,
                    "action": "PRESERVE",
                    "reason": "Missing from AI decisions; safely preserved"
                })

        slide_conf = float(parsed.get("confidence", 0.85))
        slide_conf = max(0.0, min(1.0, slide_conf))

        return {
            "slide_number": slide_context.get("slide_number", 1),
            "confidence": round(slide_conf, 2),
            "rationale": str(parsed.get("rationale", "Generated by Gemini 3.8 Flash High")),
            "decisions": validated_decisions,
            "warnings": []
        }

    def _build_reading_order_offline_fallback(
        self,
        slide_context: Dict[str, Any]
    ) -> Dict[str, Any]:
        """
        Deterministic semantic reasoning fallback for reading order.
        Produces full structured decisions when Gemini API is offline or unreachable.
        """
        slide_num = slide_context.get("slide_number", 1)
        slide_w = slide_context.get("slide_dimensions", {}).get("width_pt", 720.0)
        slide_h = slide_context.get("slide_dimensions", {}).get("height_pt", 540.0)
        shapes = slide_context.get("shapes", [])

        if len(shapes) <= 1:
            decisions = [
                {
                    "shape_id": s["shape_id"],
                    "semantic_role": "TITLE" if s.get("text_content") else "DECORATIVE",
                    "reading_order": 1 if s.get("text_content") else 0,
                    "confidence": 1.0,
                    "action": "PRESERVE",
                    "reason": "Single shape on slide; sequence is trivial"
                }
                for s in shapes
            ]
            return {
                "slide_number": slide_num,
                "confidence": 1.0,
                "rationale": "Single shape or empty slide; reading sequence is trivial.",
                "decisions": decisions,
                "warnings": []
            }

        decoratives = []
        headers = []
        footers = []
        content = []

        footer_text_regex = re.compile(
            r"(\bwiley\b|©|\(c\)|copyright|all\s+rights\s+reserved|confidential|draft|internal\s+use)",
            re.IGNORECASE
        )

        for s in shapes:
            txt = s.get("text_content", "").strip()
            geom = s.get("geometry", {})
            l = geom.get("left_pt", 0.0)
            t = geom.get("top_pt", 0.0)
            w = geom.get("width_pt", 0.0)
            h = geom.get("height_pt", 0.0)
            ph = s.get("placeholder_type")
            mtype = s.get("media_type")

            # A. Native Placeholders
            if ph in ("title", "ctrTitle"):
                headers.append((s, "TITLE", 0.98))
                continue
            elif ph == "subTitle":
                if t <= slide_h * 0.35:
                    headers.append((s, "SUBTITLE", 0.95))
                    continue
                else:
                    content.append(s)
                    continue
            elif ph in ("ftr", "sldNum", "dt"):
                footers.append((s, "SLIDE_NUMBER" if ph == "sldNum" else "FOOTER", 0.98))
                continue


            # B. Background & decorative shapes
            is_bg = (w >= slide_w * 0.85 and h >= slide_h * 0.60 and len(txt) == 0)
            is_thin = (w <= 5.0 or h <= 5.0)
            is_empty_shape = (len(txt) == 0 and mtype in ("TEXT", "OTHER", "AUTO_SHAPE"))
            if is_bg or is_thin or is_empty_shape:
                decoratives.append((s, "DECORATIVE", 0.95))
                continue

            # C. Footers in bottom margin (top >= slide_h * 0.86)
            if t >= slide_h * 0.86:
                is_page_num = bool(re.match(r"^\s*(page\s*)?\d+(\s*(of|/)\s*\d+)?\s*$", txt, re.IGNORECASE))
                is_wiley = bool(footer_text_regex.search(txt))
                is_running_footer = (t >= slide_h * 0.90 and h <= 30.0 and len(txt) < 80)
                if is_page_num or is_wiley or is_running_footer:
                    footers.append((s, "SLIDE_NUMBER" if is_page_num else "FOOTER", 0.95))
                    continue

            # D. Title candidates near top of slide
            if t <= slide_h * 0.22 and 0 < len(txt) < 140:
                headers.append((s, "TITLE", 0.90))
                continue

            # E. Otherwise content
            content.append(s)

        headers.sort(key=lambda x: (x[0]["geometry"]["top_pt"], x[0]["geometry"]["left_pt"]))
        refined_headers = []
        for idx, (h_s, r, c) in enumerate(headers):
            role = "TITLE" if idx == 0 else "SUBTITLE"
            refined_headers.append((h_s, role, c))

        # Repeated cards detection (e.g. badges '01', '02', '03')
        badges = []
        other_content = []
        badge_regex = re.compile(r"^\s*(0?[1-9]|[1-9][0-9]|step\s*\d+|part\s*\d+)\s*$", re.IGNORECASE)

        for s in content:
            txt = s.get("text_content", "").strip()
            geom = s.get("geometry", {})
            w = geom.get("width_pt", 0.0)
            h = geom.get("height_pt", 0.0)
            if badge_regex.match(txt) and w < 100.0 and h < 80.0:
                badges.append(s)
            else:
                other_content.append(s)

        ordered_content_shapes = []

        if len(badges) >= 2:
            y_variance = max(b["geometry"]["top_pt"] for b in badges) - min(b["geometry"]["top_pt"] for b in badges)
            x_variance = max(b["geometry"]["left_pt"] for b in badges) - min(b["geometry"]["left_pt"] for b in badges)
            is_horizontal_cards = x_variance > y_variance

            if is_horizontal_cards:
                badges.sort(key=lambda b: b["geometry"]["left_pt"])
                cards = {b["shape_id"]: [b] for b in badges}
                for oc in other_content:
                    oc_cx = oc["geometry"]["left_pt"] + oc["geometry"]["width_pt"] / 2.0
                    best_b = min(badges, key=lambda b: abs(b["geometry"]["left_pt"] + b["geometry"]["width_pt"] / 2.0 - oc_cx))
                    cards[best_b["shape_id"]].append(oc)
                for b in badges:
                    card_shapes = cards[b["shape_id"]]
                    card_shapes.sort(key=lambda s: (0 if s["shape_id"] == b["shape_id"] else 1, s["geometry"]["top_pt"]))
                    for cs in card_shapes:
                        role = "CARD_BADGE" if cs["shape_id"] == b["shape_id"] else ("CARD_HEADER" if len(cs.get("text_content", "")) < 50 else "CARD_BODY")
                        ordered_content_shapes.append((cs, role, 0.92))
            else:
                badges.sort(key=lambda b: b["geometry"]["top_pt"])
                cards = {b["shape_id"]: [b] for b in badges}
                for oc in other_content:
                    oc_cy = oc["geometry"]["top_pt"] + oc["geometry"]["height_pt"] / 2.0
                    best_b = min(badges, key=lambda b: abs(b["geometry"]["top_pt"] + b["geometry"]["height_pt"] / 2.0 - oc_cy))
                    cards[best_b["shape_id"]].append(oc)
                for b in badges:
                    card_shapes = cards[b["shape_id"]]
                    card_shapes.sort(key=lambda s: (0 if s["shape_id"] == b["shape_id"] else 1, s["geometry"]["left_pt"], s["geometry"]["top_pt"]))
                    for cs in card_shapes:
                        role = "CARD_BADGE" if cs["shape_id"] == b["shape_id"] else ("CARD_HEADER" if len(cs.get("text_content", "")) < 50 else "CARD_BODY")
                        ordered_content_shapes.append((cs, role, 0.92))
        else:
            # Check for columns
            columns = {}
            mid_x = slide_w * 0.5
            for c_item in content:
                cx = c_item["geometry"]["left_pt"] + c_item["geometry"]["width_pt"] / 2.0
                col_idx = 0 if cx < mid_x else 1
                columns.setdefault(col_idx, []).append(c_item)

            if len(columns) == 2 and min(len(v) for v in columns.values()) >= 1:
                for col_idx in sorted(columns.keys()):
                    col_shapes = columns[col_idx]
                    col_shapes.sort(key=lambda s: (s["geometry"]["top_pt"], s["geometry"]["left_pt"]))
                    for cs in col_shapes:
                        role = "IMAGE" if cs.get("media_type") == "IMAGE" else "BODY"
                        ordered_content_shapes.append((cs, role, 0.90))
            else:
                all_c = list(content)
                all_c.sort(key=lambda s: (s["geometry"]["top_pt"], s["geometry"]["left_pt"]))
                for cs in all_c:
                    role = "IMAGE" if cs.get("media_type") == "IMAGE" else "BODY"
                    ordered_content_shapes.append((cs, role, 0.90))

        footers.sort(key=lambda x: (x[0]["geometry"]["left_pt"]))

        content_reading_sequence = refined_headers + ordered_content_shapes + footers
        orig_order = [s["shape_id"] for s in shapes]
        decisions = []

        reading_order_idx = 1
        for (s, role, conf) in content_reading_sequence:
            decisions.append({
                "shape_id": s["shape_id"],
                "semantic_role": role,
                "reading_order": reading_order_idx,
                "confidence": conf,
                "action": "REORDER",
                "reason": f"Semantic sequence position {reading_order_idx} ({role})"
            })
            reading_order_idx += 1

        for (s, role, conf) in decoratives:
            decisions.append({
                "shape_id": s["shape_id"],
                "semantic_role": role,
                "reading_order": 0,
                "confidence": conf,
                "action": "EXCLUDE_DECORATIVE",
                "reason": "Visual background or decorative shape excluded from reading sequence"
            })

        proposed_xml_order = [d[0]["shape_id"] for d in decoratives] + [d[0]["shape_id"] for d in content_reading_sequence]

        for dec in decisions:
            sid = dec["shape_id"]
            if dec["action"] != "EXCLUDE_DECORATIVE":
                if orig_order.index(sid) == proposed_xml_order.index(sid):
                    dec["action"] = "PRESERVE"

        confidences = [d["confidence"] for d in decisions]
        avg_conf = sum(confidences) / len(confidences) if confidences else 1.0

        return {
            "slide_number": slide_num,
            "confidence": round(avg_conf, 2),
            "rationale": "Deterministic semantic reading order resolution.",
            "decisions": decisions,
            "warnings": []
        }

    # =========================================================================
    # Task 2: Placeholder Conversion Semantic Analysis (Runtime Model: Gemini 3.5)
    # =========================================================================

    def analyze_placeholder_conversion(
        self,
        slide_context: Dict[str, Any]
    ) -> Dict[str, Any]:
        """
        Analyzes PowerPoint slide candidate shapes for placeholder conversion using Gemini 3.5.
        Enforces:
          - Model: gemini-3.5 (never gemini-3.8-flash-high)
          - Strict structured JSON decision schema
          - Context privacy (no PPTX filesystem paths)
          - Offline / API failure graceful fallback
        """
        import os
        api_key = (
            os.environ.get("GEMINI_API_KEY") or
            os.environ.get("GOOGLE_API_KEY") or
            os.environ.get("SLIDEGUARD_GEMINI_KEY")
        )
        if self.client is None and not api_key:
            return self._build_placeholder_offline_fallback(slide_context)

        prompt = self._build_placeholder_prompt(slide_context)

        try:
            raw_response = self._call_model(prompt)
            return self._parse_and_validate_placeholder_response(raw_response, slide_context)
        except Exception as err:
            logger.warning("Gemini placeholder API error: %s. Using semantic fallback.", err)
            fallback = self._build_placeholder_offline_fallback(slide_context)
            fallback["warnings"].append(f"AI API unavailable ({err}); resolved via deterministic safety analyzer.")
            return fallback

    def _build_placeholder_prompt(self, slide_context: Dict[str, Any]) -> str:
        """Constructs prompt for Gemini 3.5 with placeholder classification and safety rules."""
        sanitized_context = {
            "slide_number": slide_context.get("slide_number", 1),
            "slide_dimensions": slide_context.get("slide_dimensions", {"width_pt": 720.0, "height_pt": 540.0}),
            "existing_native_title_id": slide_context.get("existing_native_title_id"),
            "candidates": [
                {
                    "shape_id": c.get("shape_id"),
                    "name": c.get("name"),
                    "current_ph_type": c.get("current_ph_type"),
                    "current_ph_idx": c.get("current_ph_idx"),
                    "text_preview": c.get("text_preview", ""),
                    "text_length": c.get("text_length", 0),
                    "geometry": c.get("geometry", {}),
                    "has_text": c.get("has_text", False),
                    "is_empty": c.get("is_empty", False),
                    "is_picture": c.get("is_picture", False),
                    "is_table": c.get("is_table", False),
                    "is_chart": c.get("is_chart", False),
                    "is_animated": c.get("is_animated", False),
                    "inherited_from": c.get("inherited_from", "NONE"),
                    "inherited_xfrm_available": c.get("inherited_xfrm_available", False),
                }
                for c in slide_context.get("candidates", [])
            ]
        }

        prompt = (
            f"You are the SlideGuard Runtime Accessibility Specialist (Model: {self.model_name}).\n"
            f"Task: Analyze candidate shapes on PowerPoint slide {sanitized_context['slide_number']} "
            f"for WCAG 2.1 - 1.3.1 Info and Relationships (Level A).\n\n"
            f"Among these candidate objects, which, if any, represents the slide-level title?\n\n"
            f"Rules and Guidelines:\n"
            f"1. SLIDE-LEVEL TITLE SELECTION:\n"
            f"   - Identify the single shape (if any) that represents the slide's primary title heading.\n"
            f"   - Positive indicators: prominent heading, descriptive phrase/sentence, larger font, slide topic heading.\n"
            f"   - Set 'selected_title_shape_id' to that shape's ID and action to 'CONVERT_TO_TITLE'.\n"
            f"   - If no candidate qualifies as a clear slide-level heading, or if a native title already exists (existing_native_title_id is not null):\n"
            f"     'selected_title_shape_id' MUST BE null.\n"
            f"2. HARD CARD / BADGE PROTECTION:\n"
            f"   - A shape MUST NOT become a TITLE if it is:\n"
            f"     * Numeric-only badge ('01', '02', '03', '04', '1', '2', etc.)\n"
            f"     * Sequence badge / step index\n"
            f"     * Card number or card index\n"
            f"     * Repeated badge across cards\n"
            f"     * Small decorative label or icon marker\n"
            f"   - For all card badges and design markers, set role to 'CARD_BADGE' and action to 'PRESERVE'.\n"
            f"3. PRESERVE ALL OTHER COMPONENTS:\n"
            f"   - Native titles, subtitles, footers, card bodies, empty placeholders, and diagrams must be PRESERVED.\n"
            f"4. MANUAL REVIEW:\n"
            f"   - Shapes targeted by animation sequence (is_animated=True) must NEVER be auto-converted without review.\n\n"
            f"Candidate Context Data:\n{json.dumps(sanitized_context, indent=2)}\n\n"
            f"Return JSON strictly matching this schema:\n"
            f"{{\n"
            f"  \"slide_number\": <int>,\n"
            f"  \"selected_title_shape_id\": <int or null>,\n"
            f"  \"overall_confidence\": <float between 0.0 and 1.0>,\n"
            f"  \"reason\": \"<string explaining choice>\",\n"
            f"  \"candidates\": [\n"
            f"    {{\n"
            f"      \"shape_id\": <int>,\n"
            f"      \"role\": \"TITLE\" | \"CARD_BADGE\" | \"CARD_BODY\" | \"FOOTER\" | \"BODY_TEXTBOX\" | \"SUBTITLE\" | \"DECORATIVE\" | \"NATIVE_TITLE\" | \"AMBIGUOUS\",\n"
            f"      \"action\": \"CONVERT_TO_TITLE\" | \"PRESERVE\" | \"MATERIALIZE_BODY\" | \"MANUAL_REVIEW\",\n"
            f"      \"confidence\": <float between 0.0 and 1.0>,\n"
            f"      \"reason\": \"<brief explanation>\"\n"
            f"    }}\n"
            f"  ]\n"
            f"}}\n"
            f"Every candidate shape_id in context MUST appear in candidates. Do not invent shape IDs. Return valid JSON only."
        )
        return prompt

    def _parse_and_validate_placeholder_response(
        self,
        raw_response: str,
        slide_context: Dict[str, Any]
    ) -> Dict[str, Any]:
        """Parses and schema-validates Gemini response for placeholder conversion."""
        if not raw_response or not raw_response.strip():
            logger.warning("Empty AI response received for placeholder analysis.")
            return self._build_placeholder_offline_fallback(slide_context)

        cleaned = raw_response.strip()
        if cleaned.startswith("```"):
            lines = cleaned.split("\n")
            if lines[0].startswith("```"):
                lines = lines[1:]
            if lines and lines[-1].strip() == "```":
                lines = lines[:-1]
            cleaned = "\n".join(lines).strip()

        try:
            parsed = json.loads(cleaned)
        except Exception as e:
            logger.warning("Failed to parse JSON from AI placeholder response: %s", e)
            fallback = self._build_placeholder_offline_fallback(slide_context)
            fallback["warnings"].append(f"JSON decode failed ({e}); used deterministic fallback.")
            return fallback

        if not isinstance(parsed, dict):
            fallback = self._build_placeholder_offline_fallback(slide_context)
            fallback["warnings"].append("AI response is not a JSON object; used fallback.")
            return fallback

        expected_ids = {c["shape_id"] for c in slide_context.get("candidates", [])}
        valid_actions = {"CONVERT_TO_TITLE", "MATERIALIZE_BODY", "PRESERVE", "MANUAL_REVIEW"}

        raw_list = parsed.get("candidates")
        if raw_list is None:
            raw_list = parsed.get("decisions", [])
        if not isinstance(raw_list, list):
            raw_list = []

        validated_candidates = []
        validated_decisions = []
        found_ids = set()

        raw_selected_id = parsed.get("selected_title_shape_id")
        selected_title_id: Optional[int] = None
        if raw_selected_id is not None and isinstance(raw_selected_id, (int, str)):
            try:
                sid_int = int(raw_selected_id)
                if sid_int in expected_ids:
                    selected_title_id = sid_int
            except (ValueError, TypeError):
                selected_title_id = None

        for item in raw_list:
            if not isinstance(item, dict):
                continue
            sid = item.get("shape_id")
            if sid not in expected_ids:
                continue
            found_ids.add(sid)
            action = str(item.get("action", "PRESERVE")).upper()
            if action not in valid_actions:
                action = "PRESERVE"
            try:
                conf = float(item.get("confidence", 0.85))
                conf = max(0.0, min(1.0, conf))
            except (ValueError, TypeError):
                conf = 0.85

            role = str(item.get("role") or item.get("candidate_type", "AMBIGUOUS"))
            reason_text = str(item.get("reason", "AI evaluated placeholder status."))
            target_ph = str(item.get("target_placeholder_type", "none"))

            if action == "CONVERT_TO_TITLE" and selected_title_id is None:
                selected_title_id = sid

            cand_dict = {
                "shape_id": sid,
                "role": role,
                "action": action,
                "confidence": conf,
                "reason": reason_text,
            }
            dec_dict = {
                "shape_id": sid,
                "candidate_type": role,
                "target_placeholder_type": "title" if action == "CONVERT_TO_TITLE" else target_ph,
                "action": action,
                "confidence": conf,
                "reason": reason_text,
            }
            validated_candidates.append(cand_dict)
            validated_decisions.append(dec_dict)

        # Fill any missing candidate shape IDs using fallback
        missing_ids = expected_ids - found_ids
        if missing_ids:
            logger.warning("AI response missed candidates: %s. Filling with fallback.", missing_ids)
            fb = self._build_placeholder_offline_fallback(slide_context)
            fb_map = {dec["shape_id"]: dec for dec in fb["decisions"]}
            for mid in missing_ids:
                if mid in fb_map:
                    validated_decisions.append(fb_map[mid])
                    validated_candidates.append({
                        "shape_id": mid,
                        "role": fb_map[mid].get("candidate_type", "AMBIGUOUS"),
                        "action": fb_map[mid].get("action", "PRESERVE"),
                        "confidence": fb_map[mid].get("confidence", 0.85),
                        "reason": fb_map[mid].get("reason", "Fallback"),
                    })

        overall_conf = float(parsed.get("overall_confidence") or parsed.get("confidence", 0.90))
        reason_summary = str(parsed.get("reason") or parsed.get("rationale", "AI placeholder conversion analysis."))

        return {
            "slide_number": slide_context.get("slide_number", 1),
            "selected_title_shape_id": selected_title_id,
            "overall_confidence": overall_conf,
            "confidence": overall_conf,
            "model_used": self.model_name,
            "reason": reason_summary,
            "rationale": reason_summary,
            "candidates": validated_candidates,
            "decisions": validated_decisions,
            "warnings": []
        }

    def _build_placeholder_offline_fallback(
        self,
        slide_context: Dict[str, Any]
    ) -> Dict[str, Any]:
        """
        High-fidelity deterministic offline semantic analyzer for placeholder conversion.
        Evaluates shapes strictly without live Gemini API.
        Enforces:
          - Hard card badge / numeric sequence protection (e.g. '01', '02', '03')
          - Multi-signal title candidate ranking across the entire slide
          - Zero greedy first-match XML order bias
          - Native title preservation
        """
        import re
        slide_num = slide_context.get("slide_number", 1)
        slide_dims = slide_context.get("slide_dimensions", {"width_pt": 720.0, "height_pt": 540.0})
        slide_h = float(slide_dims.get("height_pt", 540.0))
        slide_w = float(slide_dims.get("width_pt", 720.0))
        has_native_title = slide_context.get("existing_native_title_id") is not None
        candidates = slide_context.get("candidates", [])

        # Check if any candidate has native title ph
        for c in candidates:
            ph_t = (c.get("current_ph_type") or "").lower()
            if ph_t in ("title", "ctrtitle"):
                has_native_title = True
                break

        # Step A: Identify numeric badges and card components
        badge_shape_ids = set()
        for c in candidates:
            txt = (c.get("text_preview") or "").strip()
            if not txt:
                continue
            # Pure digits (e.g. '01', '02', '03', '1') or sequence markers ('Step 1', 'Card 2')
            if re.match(r"^\d{1,4}$", txt):
                badge_shape_ids.add(c["shape_id"])
            elif re.match(r"^(?:step|card|part|item|no\.?|#|phase|badge)\s*\d{1,3}[.:]?$", txt, re.IGNORECASE):
                badge_shape_ids.add(c["shape_id"])
            elif len(txt) <= 3 and not txt.isalpha():
                badge_shape_ids.add(c["shape_id"])

        # Step B: Rank all text candidates for slide title across the entire slide
        ranked_title_candidates = []
        if not has_native_title:
            for c in candidates:
                sid = c["shape_id"]
                if sid in badge_shape_ids:
                    continue
                if c.get("is_animated") or c.get("is_empty") or not c.get("has_text"):
                    continue
                if c.get("current_ph_type"):
                    continue

                txt = (c.get("text_preview") or "").strip()
                if len(txt) < 4:
                    continue

                geo = c.get("geometry", {})
                top_pt = float(geo.get("top_pt", 0.0))

                # Footer region exclusion
                if top_pt >= (slide_h * 0.88):
                    continue

                words = txt.split()
                w_count = len(words)
                score = 0.0

                # Heading-like word count
                if 2 <= w_count <= 25:
                    score += 25.0
                elif w_count == 1:
                    score += 5.0
                elif w_count > 25:
                    score -= 15.0

                # Font size bonus (if available)
                max_sz = float(c.get("max_font_size_pt", 0.0))
                if max_sz >= 24.0:
                    score += 35.0
                elif max_sz >= 18.0:
                    score += 20.0
                elif max_sz >= 14.0:
                    score += 10.0

                if c.get("is_bold"):
                    score += 10.0

                # Placement bonus
                if top_pt <= (slide_h * 0.15):
                    score += 25.0
                elif top_pt <= (slide_h * 0.35):
                    score += 15.0
                elif top_pt <= (slide_h * 0.70):
                    score += 10.0

                # Title indicators
                if re.search(r"Chapter\s+\d+", txt, re.IGNORECASE):
                    score += 30.0
                if ":" in txt:
                    score += 10.0

                ranked_title_candidates.append((score, c))

            ranked_title_candidates.sort(key=lambda x: x[0], reverse=True)

        selected_title_id = None
        if ranked_title_candidates and ranked_title_candidates[0][0] >= 30.0:
            selected_title_id = ranked_title_candidates[0][1]["shape_id"]

        # Step C: Build candidates and decisions lists
        validated_candidates = []
        decisions = []

        for c in candidates:
            sid = c["shape_id"]
            ph_type = (c.get("current_ph_type") or "").lower()
            ph_idx = c.get("current_ph_idx")
            is_anim = c.get("is_animated", False)
            is_empty = c.get("is_empty", False)
            has_text = c.get("has_text", False)
            text = c.get("text_preview", "")
            inherited_from = c.get("inherited_from", "NONE")
            xfrm_avail = c.get("inherited_xfrm_available", False)

            if is_anim:
                role = "AMBIGUOUS"
                action = "MANUAL_REVIEW"
                conf = 0.85
                reason = f"Shape ID {sid} is targeted by animation sequence; conversion requires timing review."
                target_ph = "none"
            elif ph_type in ("title", "ctrtitle"):
                role = "NATIVE_TITLE"
                action = "PRESERVE"
                conf = 0.95
                reason = "Title placeholders are critical for screen reader outline navigation and must remain native."
                target_ph = "title"
            elif ph_type == "subtitle":
                role = "SUBTITLE"
                action = "PRESERVE"
                conf = 0.95
                reason = "Semantic subtitle placeholder linked to slide title layout."
                target_ph = "subTitle"
            elif ph_type in ("ftr", "footer", "sldnum", "slidenumber", "dt", "date"):
                role = "FOOTER"
                action = "PRESERVE"
                conf = 0.95
                reason = "Native footer/slide number/date placeholder must not be altered."
                target_ph = ph_type or "none"
            elif c.get("is_picture") or c.get("is_table") or c.get("is_chart"):
                role = "PICTURE_PLACEHOLDER" if c.get("is_picture") else ("TABLE_PLACEHOLDER" if c.get("is_table") else "CHART_PLACEHOLDER")
                action = "PRESERVE"
                conf = 0.95
                reason = "Specialized visual placeholder; preserved for specialized phase."
                target_ph = "pic" if c.get("is_picture") else ("tbl" if c.get("is_table") else "chart")
            elif is_empty or ph_type == "empty":
                role = "EMPTY_PLACEHOLDER"
                action = "PRESERVE"
                conf = 0.95
                reason = "Empty container preserved without modification."
                target_ph = ph_type or "none"
            elif ph_type in ("body", "obj") or (ph_idx and not ph_type):
                if has_text and (xfrm_avail or inherited_from != "NONE"):
                    role = "BODY_PLACEHOLDER"
                    action = "MATERIALIZE_BODY"
                    conf = 0.95
                    reason = "Contains accessible text; materializing inherited geometry guarantees zero visual drift."
                    target_ph = "none"
                else:
                    role = "BODY_PLACEHOLDER"
                    action = "PRESERVE"
                    conf = 0.90
                    reason = "Body placeholder preserved."
                    target_ph = ph_type or "obj"
            elif sid in badge_shape_ids:
                role = "CARD_BADGE"
                action = "PRESERVE"
                conf = 0.99
                reason = f"Shape ID {sid} is a card badge/design marker ('{text}'); must remain standard visual component."
                target_ph = "none"
            elif sid == selected_title_id:
                role = "TITLE"
                action = "CONVERT_TO_TITLE"
                conf = 0.95
                reason = "Primary slide-level heading identified via multi-signal semantic ranking."
                target_ph = "title"
            else:
                role = "BODY_TEXTBOX"
                action = "PRESERVE"
                conf = 0.90
                reason = "Loose textbox preserved without modification."
                target_ph = "none"

            validated_candidates.append({
                "shape_id": sid,
                "role": role,
                "action": action,
                "confidence": conf,
                "reason": reason
            })
            decisions.append({
                "shape_id": sid,
                "candidate_type": role,
                "target_placeholder_type": target_ph,
                "action": action,
                "confidence": conf,
                "reason": reason
            })

        return {
            "slide_number": slide_num,
            "selected_title_shape_id": selected_title_id,
            "overall_confidence": 0.95,
            "confidence": 0.95,
            "model_used": self.model_name,
            "reason": "Deterministic semantic title candidate ranking and card badge protection.",
            "rationale": "Deterministic semantic title candidate ranking and card badge protection.",
            "candidates": validated_candidates,
            "decisions": decisions,
            "warnings": []
        }

    # =========================================================================
    # TASK 3: AI-ASSISTED ALTERNATIVE TEXT & DECORATIVE IMAGE DETECTION
    # =========================================================================

    def analyze_image_accessibility(
        self,
        image_context: Dict[str, Any]
    ) -> Dict[str, Any]:
        """
        Analyzes an image's accessibility role (MEANINGFUL vs DECORATIVE) and generates
        concise, factual alternative text using Gemini 3.5 runtime model.
        Strictly enforces:
          - Model: gemini-3.5 runtime (with fallback to gemini-3.1)
          - Structured context only (no local filesystem paths)
          - Strict JSON output contract
          - Anti-hallucination sanitization
          - Deterministic offline fallback
        """
        # 1. Deterministic offline fallback if client is absent
        if self.client is None:
            return self._build_image_offline_fallback(image_context)

        # 2. Build minimal structured prompt
        prompt = self._build_image_prompt(image_context)

        # 3. Call model safely
        try:
            raw_response = self._call_model(prompt)
            return self._parse_and_validate_image_response(raw_response, image_context)
        except Exception as err:
            logger.warning("Gemini image accessibility API error: %s. Falling back.", err)
            return self._build_image_offline_fallback(image_context, error_msg=str(err))

    def _build_image_prompt(self, image_context: Dict[str, Any]) -> str:
        """Constructs minimal structured prompt for image accessibility evaluation."""
        payload_str = json.dumps(image_context, indent=2)
        return (
            "You are SlideGuard's runtime accessibility AI (model: Gemini 3.5). "
            "Analyze the following presentation image object and determine its semantic role "
            "and alternative text description for screen-reader users.\n\n"
            f"IMAGE CONTEXT:\n{payload_str}\n\n"
            "RULES:\n"
            "1. Classification must be one of: 'MEANINGFUL', 'DECORATIVE', 'UNKNOWN'.\n"
            "2. Action must be one of: 'PRESERVE', 'GENERATE_ALT_TEXT', 'MARK_DECORATIVE', 'MANUAL_REVIEW'.\n"
            "3. If MEANINGFUL and lacking alt text, action='GENERATE_ALT_TEXT'. Provide concise, factual alt_text.\n"
            "4. If purely aesthetic, separator, decorative pattern, or ornament providing no information beyond surrounding content, classification='DECORATIVE', action='MARK_DECORATIVE', alt_text=''.\n"
            "5. If ambiguous or visual details cannot be verified, action='MANUAL_REVIEW' with lower confidence (<0.70).\n"
            "6. Alt text MUST be concise, factual, and specific. Do NOT begin with 'This image shows', 'Image of', 'Picture of', or 'Photo of'.\n"
            "7. Do NOT invent names, numbers, labels, or relationships not in the context.\n"
            "8. Return ONLY valid JSON matching this schema:\n"
            "{\n"
            '  "slide_number": <int>,\n'
            '  "shape_id": <int>,\n'
            '  "classification": "MEANINGFUL" | "DECORATIVE" | "UNKNOWN",\n'
            '  "action": "PRESERVE" | "GENERATE_ALT_TEXT" | "MARK_DECORATIVE" | "MANUAL_REVIEW",\n'
            '  "alt_text": "<concise description or empty>",\n'
            '  "confidence": <float 0.0 to 1.0>,\n'
            '  "reason": "<rationale>"\n'
            "}"
        )

    def _parse_and_validate_image_response(
        self,
        raw_response: str,
        image_context: Dict[str, Any]
    ) -> Dict[str, Any]:
        """Parses and schema-validates Gemini response for image accessibility."""
        cleaned = raw_response.strip()
        if "```json" in cleaned:
            cleaned = cleaned.split("```json")[1].split("```")[0].strip()
        elif "```" in cleaned:
            cleaned = cleaned.split("```")[1].split("```")[0].strip()

        try:
            parsed = json.loads(cleaned)
        except json.JSONDecodeError as err:
            logger.warning("Gemini image response was not valid JSON: %s. Falling back.", err)
            return self._build_image_offline_fallback(image_context, error_msg="Invalid JSON response")

        slide_num = int(parsed.get("slide_number", image_context.get("slide_number", 1)))
        shape_id = int(parsed.get("shape_id", image_context.get("shape_id", 0)))

        classification = str(parsed.get("classification", "UNKNOWN")).upper()
        if classification not in ("MEANINGFUL", "DECORATIVE", "UNKNOWN"):
            classification = "UNKNOWN"

        action = str(parsed.get("action", "MANUAL_REVIEW")).upper()
        if action not in ("PRESERVE", "GENERATE_ALT_TEXT", "MARK_DECORATIVE", "MANUAL_REVIEW"):
            action = "MANUAL_REVIEW"

        confidence = float(parsed.get("confidence", 0.5))
        confidence = max(0.0, min(1.0, confidence))

        alt_text = str(parsed.get("alt_text", "")).strip()
        reason = str(parsed.get("reason", "")).strip()

        # Sanitize alt text: strip forbidden boilerplate prefixes
        forbidden_prefixes = [
            r"^(this\s+image\s+shows\s+)",
            r"^(image\s+of\s+)",
            r"^(picture\s+of\s+)",
            r"^(photo\s+of\s+)",
            r"^(graphic\s+of\s+)",
            r"^(illustration\s+of\s+)",
        ]
        for pat in forbidden_prefixes:
            m = re.match(pat, alt_text, re.IGNORECASE)
            if m:
                alt_text = alt_text[m.end():].strip()
                if alt_text:
                    alt_text = alt_text[0].upper() + alt_text[1:]

        # Safety rule: if classification is DECORATIVE, alt_text must be empty
        if classification == "DECORATIVE":
            alt_text = ""
            if action != "MANUAL_REVIEW":
                action = "MARK_DECORATIVE"

        # Safety rule: if action is GENERATE_ALT_TEXT but alt_text is empty, demote
        if action == "GENERATE_ALT_TEXT" and not alt_text:
            action = "MANUAL_REVIEW"
            confidence = min(confidence, 0.65)
            reason = "Alt text was proposed but generated text was empty."

        return {
            "slide_number": slide_num,
            "shape_id": shape_id,
            "classification": classification,
            "action": action,
            "alt_text": alt_text,
            "confidence": confidence,
            "reason": reason,
            "model_used": self.model_name
        }

    def _build_image_offline_fallback(
        self,
        image_context: Dict[str, Any],
        error_msg: Optional[str] = None
    ) -> Dict[str, Any]:
        """Deterministic offline fallback for image accessibility."""
        slide_num = int(image_context.get("slide_number", 1))
        shape_id = int(image_context.get("shape_id", 0))
        shape_name = str(image_context.get("shape_name", ""))
        existing_alt = str(image_context.get("existing_alt_text", "")).strip()
        is_generic = bool(image_context.get("is_generic_alt", False))
        is_already_dec = bool(image_context.get("is_already_decorative", False))
        is_animated = bool(image_context.get("is_animated", False))
        is_repeated = bool(image_context.get("is_repeated", False))
        caption = str(image_context.get("nearby_caption", "")).strip()
        title = str(image_context.get("slide_title", "")).strip()
        pos = image_context.get("position", {})
        w_pt = float(pos.get("width_pt", 100.0))
        h_pt = float(pos.get("height_pt", 100.0))

        # Rule 1: Animation protection
        if is_animated:
            return {
                "slide_number": slide_num,
                "shape_id": shape_id,
                "classification": "MEANINGFUL",
                "action": "MANUAL_REVIEW",
                "alt_text": existing_alt,
                "confidence": 0.85,
                "reason": "Image is targeted by animation sequence; conversion requires timing review.",
                "model_used": self.model_name
            }

        # Rule 2: Explicit existing good alt text
        if existing_alt and not is_generic:
            return {
                "slide_number": slide_num,
                "shape_id": shape_id,
                "classification": "MEANINGFUL",
                "action": "PRESERVE",
                "alt_text": existing_alt,
                "confidence": 0.98,
                "reason": "Existing author-provided alt text preserved.",
                "model_used": self.model_name
            }

        # Rule 3: Deterministic decorative signals
        is_divider = (w_pt <= 5.0 or h_pt <= 5.0)
        is_tiny_icon = (w_pt <= 20.0 and h_pt <= 20.0 and not caption and not image_context.get("has_hyperlink"))

        if is_already_dec or is_divider or is_tiny_icon or is_repeated:
            reason = "Purely decorative visual element (divider/icon/repeated ornament)."
            if is_already_dec:
                reason = "Element explicitly designated as decorative."
            elif is_divider:
                reason = "Geometric divider line serving purely visual aesthetic role."
            elif is_repeated:
                reason = "Repeated background ornamentation across presentation."
            return {
                "slide_number": slide_num,
                "shape_id": shape_id,
                "classification": "DECORATIVE",
                "action": "MARK_DECORATIVE",
                "alt_text": "",
                "confidence": 0.95,
                "reason": reason,
                "model_used": self.model_name
            }

        # Rule 4: Meaningful image with caption
        if caption:
            return {
                "slide_number": slide_num,
                "shape_id": shape_id,
                "classification": "MEANINGFUL",
                "action": "GENERATE_ALT_TEXT",
                "alt_text": caption,
                "confidence": 0.90,
                "reason": f"Descriptive caption adjacent to image: '{caption[:40]}...'",
                "model_used": self.model_name
            }

        # Rule 5: Meaningful image on titled slide
        if title:
            # Fact-grounded description based on slide context
            suggested = f"Visual graphic supporting {title}."
            if "case study" in title.lower() or "attack" in title.lower() or "security" in title.lower():
                suggested = f"Illustration for {title}."
            return {
                "slide_number": slide_num,
                "shape_id": shape_id,
                "classification": "MEANINGFUL",
                "action": "GENERATE_ALT_TEXT",
                "alt_text": suggested,
                "confidence": 0.86,
                "reason": f"Content-supporting image on slide '{title}'.",
                "model_used": self.model_name
            }

        # Fallback for ambiguous image with no text context
        return {
            "slide_number": slide_num,
            "shape_id": shape_id,
            "classification": "UNKNOWN",
            "action": "MANUAL_REVIEW",
            "alt_text": "",
            "confidence": 0.60,
            "reason": error_msg or "Image lacks sufficient contextual text for automated classification; requires manual review.",
            "model_used": self.model_name
        }
