from ibm_watsonx_orchestrate.agent_builder.tools import tool
import json


@tool
def apply_brand_policy(content_json: str, policy_json: str) -> str:
    """
    Apply brand policy to presentation content.

    Args:
        content_json (str): JSON with slides/content
        policy_json (str): JSON with brand rules

    Returns:
        str: Styled presentation JSON
    """

    content = json.loads(content_json)
    policy = json.loads(policy_json)

    styled = {
        "template_path": "templates/talentia_template.pptx",
        "global_style": {
            "title_font_family": policy["fonts"]["title"],
            "body_font_family": policy["fonts"]["body"],
            "fallback_font_family": policy["fonts"]["fallback"],
            "title_font_size_pt": 28,
            "body_font_size_pt": 18,
            "subtitle_font_size_pt": 20,
            "caption_font_size_pt": 12,
            "primary_color": policy["colors"]["primary"],
            "accent_color": policy["colors"]["accent"],
            "body_color": policy["colors"]["body"],
            "logo_file": "",
            "logo_position": "top_right",
            "max_bullets_per_slide": policy["rules"]["max_bullets"]
        },
        "slides": content["slides"]
    }

    return json.dumps(styled)