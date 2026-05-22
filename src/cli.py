"""Command-line interface for PPT Skill Generator."""

from __future__ import annotations

import argparse
from pathlib import Path

from .outline_generator import generate_outline, normalize_style
from .ppt_generator import default_output_path, generate_ppt


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description="Generate a PPTX deck from structured requirements.")
    parser.add_argument("--topic", required=True, help="PPT topic, for example: AI产品发布会")
    parser.add_argument("--pages", type=int, default=8, help="Number of slides, default: 8")
    parser.add_argument("--style", default="商务风", help="商务风 / 科技风 / 教育风 / 营销风 / 极简风")
    parser.add_argument("--audience", default="通用受众", help="Target audience")
    parser.add_argument("--purpose", default="演示汇报", help="Presentation purpose")
    parser.add_argument("--requirements", default="", help="Additional content requirements")
    parser.add_argument("--output", default="", help="Output pptx path. Defaults to output/<topic>.pptx")
    return parser


def main() -> None:
    parser = build_parser()
    args = parser.parse_args()

    style = normalize_style(args.style)
    slides = generate_outline(
        topic=args.topic,
        pages=args.pages,
        style=style,
        audience=args.audience,
        purpose=args.purpose,
        requirements=args.requirements,
    )

    output = Path(args.output) if args.output else default_output_path(args.topic)
    result = generate_ppt(
        slides=slides,
        output_path=output,
        topic=args.topic,
        style=style,
        audience=args.audience,
        purpose=args.purpose,
    )

    print(f"Generated PPT: {result}")
    print(f"Slides: {len(slides)}")
    print(f"Style: {style}")
    print(f"Audience: {args.audience}")


if __name__ == "__main__":
    main()
