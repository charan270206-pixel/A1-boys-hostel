import os
import sys
from pptx import Presentation
from pptx.util import Inches

from build_presentation import create_deck
from slide_builders_part1 import build_slide_1, build_slide_2, build_slide_3, build_slide_4, build_slide_5
from slide_builders_part2 import build_slide_6, build_slide_7, build_slide_8, build_slide_9, build_slide_10
from slide_builders_part3 import build_slide_11, build_slide_12, build_slide_13, build_slide_14, build_slide_15

OUTPUT_FILENAME = "Cybersecurity_and_Cryptography_Methods.pptx"

def main():
    print("=" * 70)
    print("GENERATING CYBERSECURITY & CRYPTOGRAPHY METHODS PRESENTATION")
    print("=" * 70)

    prs = create_deck()

    slide_builders = [
        ("Slide 01: Title Slide", build_slide_1),
        ("Slide 02: Introduction to Cybersecurity", build_slide_2),
        ("Slide 03: The CIA Triad", build_slide_3),
        ("Slide 04: Common Cybersecurity Threats", build_slide_4),
        ("Slide 05: Types of Cybersecurity", build_slide_5),
        ("Slide 06: Introduction to Cryptography", build_slide_6),
        ("Slide 07: Symmetric-Key Cryptography", build_slide_7),
        ("Slide 08: Asymmetric-Key Cryptography", build_slide_8),
        ("Slide 09: Hash Functions & Digital Signatures", build_slide_9),
        ("Slide 10: Cryptographic Protocols in Modern Networks", build_slide_10),
        ("Slide 11: Authentication & Access Control (Zero Trust)", build_slide_11),
        ("Slide 12: Real-World Cyber Attacks & Case Studies", build_slide_12),
        ("Slide 13: Emerging Frontiers: Quantum & AI in Security", build_slide_13),
        ("Slide 14: Defense-in-Depth & Best Practices", build_slide_14),
        ("Slide 15: Conclusion, Takeaways & Seminar Discussion", build_slide_15)
    ]

    for title, builder in slide_builders:
        print(f"Building {title}...")
        builder(prs)

    # Save presentation
    prs.save(OUTPUT_FILENAME)
    print("=" * 70)
    print(f"SUCCESS: Saved presentation to '{OUTPUT_FILENAME}'")
    print("=" * 70)

    # Verification checks
    verify_prs = Presentation(OUTPUT_FILENAME)
    num_slides = len(verify_prs.slides)
    file_size_kb = os.path.getsize(OUTPUT_FILENAME) / 1024
    
    print(f"• Total Slides Generated: {num_slides} (Target: 15)")
    print(f"• File Size on Disk:      {file_size_kb:.2f} KB")
    print(f"• Slide Width:            {verify_prs.slide_width.inches:.3f} inches (16:9 widescreen)")
    print(f"• Slide Height:           {verify_prs.slide_height.inches:.3f} inches (16:9 widescreen)")

    assert num_slides == 15, f"Expected 15 slides, got {num_slides}"

    # Verify notes and shapes on every slide
    all_notes_present = True
    for idx, slide in enumerate(verify_prs.slides, 1):
        num_shapes = len(slide.shapes)
        has_notes = bool(slide.notes_slide.notes_text_frame.text.strip())
        notes_len = len(slide.notes_slide.notes_text_frame.text.strip())
        if not has_notes:
            all_notes_present = False
            print(f"  [!] Warning: Slide {idx} is missing speaker notes!")
        else:
            print(f"  [OK] Slide {idx:02d}: {num_shapes} native shapes/objects, {notes_len} chars of speaker notes.")

    if all_notes_present:
        print("[OK] ALL 15 SLIDES HAVE COMPREHENSIVE SPEAKER NOTES & NATIVE POWERPOINT OBJECTS.")
    else:
        print("[!] Some slides were missing notes.")

if __name__ == '__main__':
    main()
