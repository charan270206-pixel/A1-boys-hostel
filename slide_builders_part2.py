from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.enum.shapes import MSO_SHAPE
from build_presentation import (
    apply_background, add_header, add_card, add_photo_card, set_speaker_notes,
    COLOR_BG_DARK, COLOR_CARD_BG, COLOR_CARD_BORDER, COLOR_CARD_HIGHLIGHT, COLOR_INNER_BOX,
    COLOR_CYAN_ACCENT, COLOR_CYAN_LIGHT, COLOR_BLUE_ACCENT, COLOR_BLUE_DARK,
    COLOR_GREEN_ACCENT, COLOR_RED_ACCENT, COLOR_AMBER_ACCENT, COLOR_PURPLE_ACCENT,
    COLOR_TEXT_PRIMARY, COLOR_TEXT_SECONDARY, COLOR_TEXT_MUTED,
    FONT_HEADING, FONT_BODY, FONT_MONO
)

# ==============================================================================
# SLIDE 6: INTRODUCTION TO CRYPTOGRAPHY
# ==============================================================================
def build_slide_6(prs):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    apply_background(slide)
    add_header(
        slide,
        badge_text="MODULE 02: CRYPTOGRAPHIC PRIMITIVES",
        title_text="Introduction to Cryptography: The Science of Secrets",
        subtitle_text="Transforming intelligible information into pseudorandom ciphertext through reversible mathematical transformations",
        slide_num=6
    )

    # ---------------------------------------------------------
    # TOP SECTION: NATIVE EDITABLE FLOWCHART
    # Flow: Plaintext (P) -> [Key K] -> Encryption -> Ciphertext (C) -> [Key K] -> Decryption -> Plaintext (P)
    # ---------------------------------------------------------
    flow_y = Inches(1.65)
    flow_h = Inches(1.25)
    flow_card = add_card(slide, Inches(0.8), flow_y, Inches(11.733), flow_h, bg_color=COLOR_CARD_HIGHLIGHT, border_color=COLOR_CYAN_ACCENT)

    # Flowchart title inside container
    f_tb = slide.shapes.add_textbox(Inches(0.95), flow_y + Inches(0.04), Inches(11.4), Inches(0.24))
    f_p = f_tb.text_frame.paragraphs[0]
    f_p.text = "CANONICAL CRYPTOGRAPHIC FLOW: MATHEMATICAL TRANSFORMATION PIPELINE"
    f_p.font.name = FONT_HEADING
    f_p.font.size = Pt(8.5)
    f_p.font.bold = True
    f_p.font.color.rgb = COLOR_CYAN_LIGHT

    # Diagram Elements:
    # 1. Plaintext Box
    box_p1 = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(1.0), flow_y + Inches(0.32), Inches(1.8), Inches(0.75))
    box_p1.fill.solid()
    box_p1.fill.fore_color.rgb = COLOR_INNER_BOX
    box_p1.line.color.rgb = COLOR_BLUE_ACCENT
    box_p1.line.width = Pt(1.5)
    tf1 = box_p1.text_frame
    tf1.word_wrap = True
    p = tf1.paragraphs[0]
    p.text = "PLAINTEXT (P)"
    p.alignment = PP_ALIGN.CENTER
    p.font.name = FONT_MONO
    p.font.size = Pt(9.5)
    p.font.bold = True
    p.font.color.rgb = COLOR_TEXT_PRIMARY
    p_sub = tf1.add_paragraph()
    p_sub.text = '"Confidential Data"'
    p_sub.alignment = PP_ALIGN.CENTER
    p_sub.font.size = Pt(8)
    p_sub.font.color.rgb = COLOR_CYAN_LIGHT

    # Arrow 1
    arr1 = slide.shapes.add_shape(MSO_SHAPE.RIGHT_ARROW, Inches(2.9), flow_y + Inches(0.55), Inches(0.45), Inches(0.3))
    arr1.fill.solid()
    arr1.fill.fore_color.rgb = COLOR_CYAN_ACCENT
    arr1.line.fill.background()

    # 2. Encryption Function + Key
    box_enc = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(3.45), flow_y + Inches(0.32), Inches(2.1), Inches(0.75))
    box_enc.fill.solid()
    box_enc.fill.fore_color.rgb = COLOR_INNER_BOX
    box_enc.line.color.rgb = COLOR_CYAN_ACCENT
    box_enc.line.width = Pt(1.5)
    tf2 = box_enc.text_frame
    tf2.word_wrap = True
    p = tf2.paragraphs[0]
    p.text = "ENCRYPTION: E_K"
    p.alignment = PP_ALIGN.CENTER
    p.font.name = FONT_MONO
    p.font.size = Pt(9.5)
    p.font.bold = True
    p.font.color.rgb = COLOR_CYAN_LIGHT
    p_sub = tf2.add_paragraph()
    p_sub.text = "+ Cipher Key (K)"
    p_sub.alignment = PP_ALIGN.CENTER
    p_sub.font.size = Pt(8)
    p_sub.font.color.rgb = COLOR_GREEN_ACCENT

    # Arrow 2
    arr2 = slide.shapes.add_shape(MSO_SHAPE.RIGHT_ARROW, Inches(5.65), flow_y + Inches(0.55), Inches(0.45), Inches(0.3))
    arr2.fill.solid()
    arr2.fill.fore_color.rgb = COLOR_CYAN_ACCENT
    arr2.line.fill.background()

    # 3. Ciphertext Box
    box_c = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(6.2), flow_y + Inches(0.32), Inches(2.0), Inches(0.75))
    box_c.fill.solid()
    box_c.fill.fore_color.rgb = COLOR_INNER_BOX
    box_c.line.color.rgb = COLOR_RED_ACCENT
    box_c.line.width = Pt(1.5)
    tf3 = box_c.text_frame
    tf3.word_wrap = True
    p = tf3.paragraphs[0]
    p.text = "CIPHERTEXT (C)"
    p.alignment = PP_ALIGN.CENTER
    p.font.name = FONT_MONO
    p.font.size = Pt(9.5)
    p.font.bold = True
    p.font.color.rgb = COLOR_RED_ACCENT
    p_sub = tf3.add_paragraph()
    p_sub.text = '7f8a9e4b2d1c...'
    p_sub.alignment = PP_ALIGN.CENTER
    p_sub.font.size = Pt(8)
    p_sub.font.color.rgb = COLOR_TEXT_MUTED

    # Arrow 3
    arr3 = slide.shapes.add_shape(MSO_SHAPE.RIGHT_ARROW, Inches(8.3), flow_y + Inches(0.55), Inches(0.45), Inches(0.3))
    arr3.fill.solid()
    arr3.fill.fore_color.rgb = COLOR_CYAN_ACCENT
    arr3.line.fill.background()

    # 4. Decryption Function + Key
    box_dec = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(8.85), flow_y + Inches(0.32), Inches(2.1), Inches(0.75))
    box_dec.fill.solid()
    box_dec.fill.fore_color.rgb = COLOR_INNER_BOX
    box_dec.line.color.rgb = COLOR_CYAN_ACCENT
    box_dec.line.width = Pt(1.5)
    tf4 = box_dec.text_frame
    tf4.word_wrap = True
    p = tf4.paragraphs[0]
    p.text = "DECRYPTION: D_K"
    p.alignment = PP_ALIGN.CENTER
    p.font.name = FONT_MONO
    p.font.size = Pt(9.5)
    p.font.bold = True
    p.font.color.rgb = COLOR_CYAN_LIGHT
    p_sub = tf4.add_paragraph()
    p_sub.text = "+ Cipher Key (K)"
    p_sub.alignment = PP_ALIGN.CENTER
    p_sub.font.size = Pt(8)
    p_sub.font.color.rgb = COLOR_GREEN_ACCENT

    # Arrow 4
    arr4 = slide.shapes.add_shape(MSO_SHAPE.RIGHT_ARROW, Inches(11.05), flow_y + Inches(0.55), Inches(0.4), Inches(0.3))
    arr4.fill.solid()
    arr4.fill.fore_color.rgb = COLOR_CYAN_ACCENT
    arr4.line.fill.background()

    # 5. Restored Plaintext Box (Right edge)
    box_p2 = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(11.5), flow_y + Inches(0.32), Inches(0.9), Inches(0.75))
    box_p2.fill.solid()
    box_p2.fill.fore_color.rgb = COLOR_INNER_BOX
    box_p2.line.color.rgb = COLOR_GREEN_ACCENT
    box_p2.line.width = Pt(1.5)
    tf5 = box_p2.text_frame
    tf5.word_wrap = True
    p = tf5.paragraphs[0]
    p.text = "PLAINTEXT"
    p.alignment = PP_ALIGN.CENTER
    p.font.name = FONT_MONO
    p.font.size = Pt(7.5)
    p.font.bold = True
    p.font.color.rgb = COLOR_GREEN_ACCENT

    # ---------------------------------------------------------
    # BOTTOM SECTION: 4 CONCEPT CARDS & RIGHT PHOTO
    # ---------------------------------------------------------
    left_w = Inches(7.3)
    c_y = Inches(3.05)
    c_h = Inches(3.75)
    card_terms = add_card(slide, Inches(0.8), c_y, left_w, c_h, border_color=COLOR_CARD_BORDER)

    tb_t = slide.shapes.add_textbox(Inches(0.95), c_y + Inches(0.1), left_w - Inches(0.3), c_h - Inches(0.2))
    tf_t = tb_t.text_frame
    tf_t.word_wrap = True

    p_th = tf_t.paragraphs[0]
    p_th.text = "FUNDAMENTAL CRYPTOGRAPHIC TAXONOMY & PRINCIPLES"
    p_th.font.name = FONT_HEADING
    p_th.font.size = Pt(11)
    p_th.font.bold = True
    p_th.font.color.rgb = COLOR_CYAN_LIGHT

    terms = [
        ("Plaintext (P):", "The unencrypted, human-readable or system-parseable original data stream before applying any cipher functions."),
        ("Ciphertext (C):", "The unreadable, pseudorandom output produced by an encryption algorithm. To an adversary without the key, it appears as pure entropy."),
        ("Cipher (Algorithm):", "The mathematical function or state machine executing cryptographic transformations. Expressed formally as C = E_K(P) and P = D_K(C)."),
        ("Cryptographic Key (K):", "The secret parameter (bit string) that governs the cipher's behavior. The strength of the system depends on key length and unpredictability."),
        ("Cryptanalysis vs Cryptology:", "Cryptography is the science of designing ciphers; Cryptanalysis is the science of breaking ciphers. Together they constitute Cryptology.")
    ]
    for lbl, desc in terms:
        p = tf_t.add_paragraph()
        p.space_before = Pt(3)
        r1 = p.add_run()
        r1.text = f"• {lbl} "
        r1.font.bold = True
        r1.font.size = Pt(9)
        r1.font.color.rgb = COLOR_BLUE_ACCENT
        r2 = p.add_run()
        r2.text = desc
        r2.font.size = Pt(9)
        r2.font.color.rgb = COLOR_TEXT_SECONDARY

    # Kerckhoffs's Principle Highlight Inside Terms Card
    p_k = tf_t.add_paragraph()
    p_k.space_before = Pt(6)
    r_kh = p_k.add_run()
    r_kh.text = "KERCKHOFFS'S PRINCIPLE (1883): "
    r_kh.font.bold = True
    r_kh.font.size = Pt(9)
    r_kh.font.color.rgb = COLOR_AMBER_ACCENT
    r_kd = p_k.add_run()
    r_kd.text = '"A cryptographic system should remain secure even if everything about the system—including algorithm mathematical design—is public knowledge, as long as the key remains secret." (Zero Security Through Obscurity!)'
    r_kd.font.size = Pt(9)
    r_kd.font.color.rgb = COLOR_TEXT_PRIMARY

    # Right Column: Photo & Cryptographic Math Callout
    rx = Inches(8.4)
    rw = Inches(4.133)
    
    add_photo_card(
        slide,
        'assets/slide6_crypto.jpg',
        rx, c_y, rw, Inches(2.45),
        caption="Figure 6.1: Cryptographic cipher lock securing high-entropy data pipelines.",
        border_color=COLOR_CYAN_ACCENT
    )

    math_card = add_card(slide, rx, c_y + Inches(2.9), rw, Inches(0.85), bg_color=COLOR_INNER_BOX, border_color=COLOR_BLUE_ACCENT)
    tb_m = slide.shapes.add_textbox(rx + Inches(0.1), c_y + Inches(2.95), rw - Inches(0.2), Inches(0.75))
    tf_m = tb_m.text_frame
    tf_m.word_wrap = True
    p_mh = tf_m.paragraphs[0]
    p_mh.text = "MATHEMATICAL FORMULATION"
    p_mh.font.name = FONT_MONO
    p_mh.font.size = Pt(8.5)
    p_mh.font.bold = True
    p_mh.font.color.rgb = COLOR_CYAN_LIGHT

    p_md = tf_m.add_paragraph()
    p_md.space_before = Pt(2)
    p_md.text = "Encryption:  C = E_K(P)\nDecryption:  P = D_K(C) = D_K(E_K(P))"
    p_md.font.name = FONT_MONO
    p_md.font.size = Pt(9)
    p_md.font.bold = True
    p_md.font.color.rgb = COLOR_TEXT_PRIMARY

    # Speaker Notes
    set_speaker_notes(slide, """
SLIDE 6: INTRODUCTION TO CRYPTOGRAPHY (SPEAKER SCRIPT)

• CORE TALKING POINTS:
1. "Now we enter the heart of digital protection: Cryptography. Derived from the Greek words 'kryptos' meaning hidden and 'graphein' meaning to write."
2. "Let's walk through the flowchart at the top:
   - We start with Plaintext (P), which is unencrypted readable data.
   - We pass it through an Encryption Algorithm (E) along with a mathematical Key (K).
   - The output is Ciphertext (C)—a seemingly random, high-entropy sequence of bytes.
   - At the destination, the Decryption Algorithm (D) uses Key (K) to invert the mathematical transformation and recover the original Plaintext (P)."
3. "Pay special attention to Kerckhoffs's Principle at the bottom: This is the golden rule of modern security. Never invent your own secret proprietary algorithm and assume you are safe because the code is secret. That is called 'Security through Obscurity', and it always fails. In modern cryptography, the algorithms—like AES, RSA, and SHA-256—are completely open source and scrutinized by mathematicians worldwide. The entire security rests solely on keeping the KEY secret."
4. "Claude Shannon formalized this as Shannon's Maxim: 'The enemy knows the system'. If your algorithm leaks, your system must remain unbroken as long as the key is safe."

• TRANSITION:
"Now, how many keys do we use? If we use one single shared key, that is Symmetric Cryptography. If we use a mathematical pair of keys, that is Asymmetric Cryptography. Let us explore Symmetric Ciphers next."
""")


# ==============================================================================
# SLIDE 7: SYMMETRIC-KEY CRYPTOGRAPHY
# ==============================================================================
def build_slide_7(prs):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    apply_background(slide)
    add_header(
        slide,
        badge_text="MODULE 02: SECRET-KEY CIPHERS",
        title_text="Symmetric Cryptography: High-Speed Shared Secrets",
        subtitle_text="Single-key cryptography engineered for ultra-high throughput bulk data encryption across disks and channels",
        slide_num=7
    )

    # ---------------------------------------------------------
    # LEFT & CENTER: ARCHITECTURE, ALGORITHMS & CHALLENGES
    # ---------------------------------------------------------
    left_w = Inches(7.3)
    
    # Card 1: Core Mechanics & Architecture
    c1 = add_card(slide, Inches(0.8), Inches(1.65), left_w, Inches(1.85), border_color=COLOR_CYAN_ACCENT)
    tb1 = slide.shapes.add_textbox(Inches(0.95), Inches(1.72), left_w - Inches(0.3), Inches(1.7))
    tf1 = tb1.text_frame
    tf1.word_wrap = True
    p1 = tf1.paragraphs[0]
    p1.text = "1. Core Principles of Symmetric-Key Ciphers"
    p1.font.name = FONT_HEADING
    p1.font.size = Pt(11)
    p1.font.bold = True
    p1.font.color.rgb = COLOR_CYAN_LIGHT

    bullets1 = [
        ("Identical Shared Secret:", "Sender and recipient share an exact identical secret key (K) for both encryption and decryption. E_K(P) = C and D_K(C) = P."),
        ("Computational Efficiency:", "Operates on direct bitwise substitutions (S-Boxes), permutations (P-Boxes), and XOR operations. Extremely fast and hardware-accelerated (AES-NI instruction set on modern Intel/AMD/ARM CPUs)."),
        ("Primary Use Case:", "Ideal for bulk encryption of large datasets: solid-state drive encryption (BitLocker, LUKS), high-speed network pipes (IPsec VPNs), and active TLS session records.")
    ]
    for lbl, desc in bullets1:
        p = tf1.add_paragraph()
        p.space_before = Pt(2)
        r1 = p.add_run()
        r1.text = f"• {lbl} "
        r1.font.bold = True
        r1.font.size = Pt(9)
        r1.font.color.rgb = COLOR_BLUE_ACCENT
        r2 = p.add_run()
        r2.text = desc
        r2.font.size = Pt(9)
        r2.font.color.rgb = COLOR_TEXT_SECONDARY

    # Card 2: Algorithms Comparison & Block vs Stream
    c2 = add_card(slide, Inches(0.8), Inches(3.6), left_w, Inches(2.0), border_color=COLOR_CARD_BORDER)
    tb2 = slide.shapes.add_textbox(Inches(0.95), Inches(3.67), left_w - Inches(0.3), Inches(1.85))
    tf2 = tb2.text_frame
    tf2.word_wrap = True
    p2 = tf2.paragraphs[0]
    p2.text = "2. Prominent Symmetric Algorithms & Modes"
    p2.font.name = FONT_HEADING
    p2.font.size = Pt(11)
    p2.font.bold = True
    p2.font.color.rgb = COLOR_BLUE_ACCENT

    algos = [
        ("AES (Advanced Encryption Standard - FIPS 197):", "Rijndael cipher utilizing a Substitution-Permutation Network (SPN). Supports 128, 192, and 256-bit keys with 10, 12, or 14 transformation rounds. The universal global gold standard; brute force is mathematically impossible with current computing."),
        ("ChaCha20-Poly1305:", "High-speed 256-bit stream cipher paired with Poly1305 authenticator (AEAD). Outperforms AES on mobile devices without dedicated cryptographic hardware. Standard in WireGuard and TLS 1.3."),
        ("DES & 3DES (Legacy):", "Data Encryption Standard (56-bit key) developed in 1977. Deprecated in 2005 due to brute-force vulnerability (cracked in under 24 hours). Triple-DES (3DES) applies 3 passes but is being decommissioned by NIST.")
    ]
    for lbl, desc in algos:
        p = tf2.add_paragraph()
        p.space_before = Pt(3)
        r1 = p.add_run()
        r1.text = f"• {lbl} "
        r1.font.bold = True
        r1.font.size = Pt(8.8)
        r1.font.color.rgb = COLOR_CYAN_LIGHT
        r2 = p.add_run()
        r2.text = desc
        r2.font.size = Pt(8.8)
        r2.font.color.rgb = COLOR_TEXT_SECONDARY

    # Card 3: The Key Distribution Dilemma (Bottom Banner across left)
    c3 = add_card(slide, Inches(0.8), Inches(5.7), left_w, Inches(1.15), bg_color=COLOR_CARD_HIGHLIGHT, border_color=COLOR_RED_ACCENT)
    tb3 = slide.shapes.add_textbox(Inches(0.95), Inches(5.75), left_w - Inches(0.3), Inches(1.05))
    tf3 = tb3.text_frame
    tf3.word_wrap = True
    p3 = tf3.paragraphs[0]
    p3.text = "3. THE CRITICAL VULNERABILITY: KEY DISTRIBUTION PROBLEM"
    p3.font.name = FONT_HEADING
    p3.font.size = Pt(9.5)
    p3.font.bold = True
    p3.font.color.rgb = COLOR_RED_ACCENT

    kd_text = (
        "• Scalability Bottleneck: For N users to communicate securely pairwise, exactly N(N-1)/2 unique shared keys are required! For 1,000 users, that requires 499,500 distinct secret keys.\n"
        "• The Insecure Medium Dilemma: How do two parties across the globe share key K securely without an attacker intercepting it in transit? This challenge necessitated Asymmetric Cryptography."
    )
    p3_d = tf3.add_paragraph()
    p3_d.space_before = Pt(2)
    p3_d.text = kd_text
    p3_d.font.name = FONT_BODY
    p3_d.font.size = Pt(8.5)
    p3_d.font.color.rgb = COLOR_TEXT_PRIMARY

    # ---------------------------------------------------------
    # RIGHT COLUMN: HARDWARE CRYPTO PHOTO & KEYSPACE METRIC
    # ---------------------------------------------------------
    rx = Inches(8.4)
    rw = Inches(4.133)
    
    add_photo_card(
        slide,
        'assets/slide7_symmetric.jpg',
        rx, Inches(1.65), rw, Inches(3.4),
        caption="Figure 7.1: Cryptographic microcontrollers and HSMs generating symmetric keys.",
        border_color=COLOR_CYAN_ACCENT
    )

    metric_card = add_card(slide, rx, Inches(5.45), rw, Inches(1.4), bg_color=COLOR_INNER_BOX, border_color=COLOR_GREEN_ACCENT)
    tb_m = slide.shapes.add_textbox(rx + Inches(0.12), Inches(5.52), rw - Inches(0.24), Inches(1.25))
    tf_m = tb_m.text_frame
    tf_m.word_wrap = True
    p_mh = tf_m.paragraphs[0]
    p_mh.text = "AES-256 KEYSPACE COMPUTATIONAL STRENGTH"
    p_mh.font.name = FONT_MONO
    p_mh.font.size = Pt(9)
    p_mh.font.bold = True
    p_mh.font.color.rgb = COLOR_GREEN_ACCENT

    m_desc = [
        "Keyspace: 2^256 possible key combinations ≈ 1.15 × 10^77 keys.",
        "Brute Force Impossibility: Even if all the supercomputers on Earth checked 1 trillion keys per second, it would take more than the age of the universe to exhaust the keyspace."
    ]
    for md in m_desc:
        p = tf_m.add_paragraph()
        p.space_before = Pt(2)
        p.text = f"✔ {md}"
        p.font.name = FONT_BODY
        p.font.size = Pt(8.2)
        p.font.color.rgb = COLOR_TEXT_SECONDARY

    # Speaker Notes
    set_speaker_notes(slide, """
SLIDE 7: SYMMETRIC-KEY CRYPTOGRAPHY (SPEAKER SCRIPT)

• CORE TALKING POINTS:
1. "In Symmetric Cryptography, Alice and Bob use the EXACT same secret key to both lock and unlock the digital vault. Think of it like a physical bank deposit safe with two identical physical keys."
2. "Why do we love symmetric ciphers? SPEED! Modern processors from Intel and AMD have dedicated AES-NI silicon instructions baked into the CPU die. AES can encrypt data at speeds exceeding 5 to 10 Gigabytes per second with virtually zero CPU overhead."
3. "Notice AES-256: It runs 14 mathematical rounds of byte substitution, row shifting, column mixing, and round key XORing. As shown in the bottom right metric, cracking AES-256 by brute force would take longer than the age of our universe."
4. "However, look at the red box on the bottom left: The Key Distribution Problem!
   - Formula: N(N - 1) / 2. If a company has 1,000 employees who need to talk privately to each other, they need nearly 500,000 unique keys!
   - Worse: How does Alice send Bob the key over the internet in the first place? If she emails it, an attacker intercepts it. If she travels to Bob with a USB flash drive, it doesn't scale to modern cloud computing."
5. "This critical dilemma remained unsolved for thousands of years until 1976, when public-key cryptography was born."

• TRANSITION:
"Let us examine how Asymmetric Cryptography solved the key distribution problem once and for all."
""")


# ==============================================================================
# SLIDE 8: ASYMMETRIC-KEY CRYPTOGRAPHY
# ==============================================================================
def build_slide_8(prs):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    apply_background(slide)
    add_header(
        slide,
        badge_text="MODULE 02: PUBLIC-KEY CRYPTOGRAPHY",
        title_text="Asymmetric Cryptography: The Dual-Key Revolution",
        subtitle_text="Solving the key distribution dilemma through mathematically coupled public and private key pairs",
        slide_num=8
    )

    # ---------------------------------------------------------
    # LEFT SECTION: DUAL-KEY CONCEPT & CORE ALGORITHMS
    # ---------------------------------------------------------
    left_w = Inches(7.3)
    
    # Card 1: The Dual-Key Trapdoor Mechanics
    c1 = add_card(slide, Inches(0.8), Inches(1.65), left_w, Inches(1.85), border_color=COLOR_CYAN_ACCENT)
    tb1 = slide.shapes.add_textbox(Inches(0.95), Inches(1.72), left_w - Inches(0.3), Inches(1.7))
    tf1 = tb1.text_frame
    tf1.word_wrap = True
    p1 = tf1.paragraphs[0]
    p1.text = "1. The Public-Private Key Pair Mechanics"
    p1.font.name = FONT_HEADING
    p1.font.size = Pt(11)
    p1.font.bold = True
    p1.font.color.rgb = COLOR_CYAN_LIGHT

    bullets1 = [
        ("Asymmetric Key Pair:", "Each participant generates two mathematically linked keys: a Public Key (freely distributed to the world) and a Private Key (kept strictly confidential and never shared)."),
        ("One-Way Trapdoor Functions:", "Easy to compute in the forward direction, but computationally infeasible to invert without special trapdoor knowledge (the private key)."),
        ("Confidentiality Flow:", "Alice encrypts message P using Bob's Public Key: C = E_Kpub(P). Once encrypted, ONLY Bob's Private Key can decrypt it: P = D_Kpriv(C). Even Alice cannot decrypt the ciphertext she just created!")
    ]
    for lbl, desc in bullets1:
        p = tf1.add_paragraph()
        p.space_before = Pt(2)
        r1 = p.add_run()
        r1.text = f"• {lbl} "
        r1.font.bold = True
        r1.font.size = Pt(9)
        r1.font.color.rgb = COLOR_BLUE_ACCENT
        r2 = p.add_run()
        r2.text = desc
        r2.font.size = Pt(9)
        r2.font.color.rgb = COLOR_TEXT_SECONDARY

    # Card 2: Prominent Asymmetric Algorithms
    c2 = add_card(slide, Inches(0.8), Inches(3.6), left_w, Inches(1.95), border_color=COLOR_CARD_BORDER)
    tb2 = slide.shapes.add_textbox(Inches(0.95), Inches(3.67), left_w - Inches(0.3), Inches(1.8))
    tf2 = tb2.text_frame
    tf2.word_wrap = True
    p2 = tf2.paragraphs[0]
    p2.text = "2. Prominent Asymmetric Algorithms: RSA, ECC & Diffie-Hellman"
    p2.font.name = FONT_HEADING
    p2.font.size = Pt(11)
    p2.font.bold = True
    p2.font.color.rgb = COLOR_BLUE_ACCENT

    algos2 = [
        ("RSA (Rivest-Shamir-Adleman - 1977):", "Relies on the computational hardness of factoring the product of two extremely large prime numbers (n = p × q). Standard key lengths: 2048 to 4096 bits. Widely used for digital certificates and legacy web security."),
        ("ECC (Elliptic Curve Cryptography):", "Relies on the discrete logarithm problem over elliptic curve point groups (y^2 = x^3 + ax + b). Exceptional efficiency: a 256-bit ECC key yields security equivalent to a 3072-bit RSA key with 90% less energy and bandwidth!"),
        ("Diffie-Hellman Key Exchange (DH / ECDH):", "Groundbreaking protocol enabling two remote parties to negotiate a shared symmetric key over an entirely public, eavesdropped channel without transmitting the key itself.")
    ]
    for lbl, desc in algos2:
        p = tf2.add_paragraph()
        p.space_before = Pt(3)
        r1 = p.add_run()
        r1.text = f"• {lbl} "
        r1.font.bold = True
        r1.font.size = Pt(8.8)
        r1.font.color.rgb = COLOR_CYAN_LIGHT
        r2 = p.add_run()
        r2.text = desc
        r2.font.size = Pt(8.8)
        r2.font.color.rgb = COLOR_TEXT_SECONDARY

    # Card 3: Hybrid Cryptosystem (Bottom left banner)
    c3 = add_card(slide, Inches(0.8), Inches(5.65), left_w, Inches(1.2), bg_color=COLOR_CARD_HIGHLIGHT, border_color=COLOR_GREEN_ACCENT)
    tb3 = slide.shapes.add_textbox(Inches(0.95), Inches(5.72), left_w - Inches(0.3), Inches(1.05))
    tf3 = tb3.text_frame
    tf3.word_wrap = True
    p3 = tf3.paragraphs[0]
    p3.text = "3. MODERN INDUSTRY REALITY: THE HYBRID ENCRYPTION PARADIGM"
    p3.font.name = FONT_HEADING
    p3.font.size = Pt(9.5)
    p3.font.bold = True
    p3.font.color.rgb = COLOR_GREEN_ACCENT

    hyb_text = (
        "• Best of Both Worlds: Asymmetric ciphers are ~1,000x computationally slower than symmetric ciphers. Therefore, modern systems (TLS, SSH, PGP) NEVER encrypt bulk files with RSA/ECC.\n"
        "• Practical Implementation: Asymmetric cryptography is used ONLY during the initial handshake to securely negotiate an ephemeral symmetric session key (AES-256), which then encrypts the actual bulk data stream at wire speed!"
    )
    p3_d = tf3.add_paragraph()
    p3_d.space_before = Pt(2)
    p3_d.text = hyb_text
    p3_d.font.name = FONT_BODY
    p3_d.font.size = Pt(8.5)
    p3_d.font.color.rgb = COLOR_TEXT_PRIMARY

    # ---------------------------------------------------------
    # RIGHT COLUMN: PHOTO & COMPARISON TABLE
    # ---------------------------------------------------------
    rx = Inches(8.4)
    rw = Inches(4.133)
    
    add_photo_card(
        slide,
        'assets/slide8_asymmetric.jpg',
        rx, Inches(1.65), rw, Inches(2.2),
        caption="Figure 8.1: Mathematically bound public-private key cryptography model.",
        border_color=COLOR_CYAN_ACCENT
    )

    # Native PowerPoint Table: Symmetric vs Asymmetric Comparison
    table_shape = slide.shapes.add_table(5, 3, rx, Inches(4.25), rw, Inches(2.6))
    table = table_shape.table
    table.columns[0].width = Inches(1.3)
    table.columns[1].width = Inches(1.4)
    table.columns[2].width = Inches(1.433)

    headers = ["Feature", "Symmetric", "Asymmetric"]
    for c_idx, h_text in enumerate(headers):
        cell = table.cell(0, c_idx)
        cell.fill.solid()
        cell.fill.fore_color.rgb = COLOR_INNER_BOX
        p = cell.text_frame.paragraphs[0]
        p.text = h_text
        p.font.name = FONT_HEADING
        p.font.size = Pt(8.5)
        p.font.bold = True
        p.font.color.rgb = COLOR_CYAN_LIGHT

    rows_data = [
        ("Key Pair", "1 Shared Secret Key", "2 Keys (Public + Private)"),
        ("Execution Speed", "Ultra-fast (Gbps throughput)", "Computationally slow (~1000x slower)"),
        ("Key Distribution", "Highly vulnerable O(N^2)", "Completely solved O(N)"),
        ("Primary Domain", "Bulk Data, Disks, Tunnels", "Key Exchange, Signatures, Handshakes")
    ]
    for r_idx, row_values in enumerate(rows_data):
        for c_idx, val in enumerate(row_values):
            cell = table.cell(r_idx + 1, c_idx)
            cell.fill.solid()
            cell.fill.fore_color.rgb = COLOR_CARD_BG if r_idx % 2 == 0 else COLOR_INNER_BOX
            p = cell.text_frame.paragraphs[0]
            p.text = val
            p.font.name = FONT_BODY
            p.font.size = Pt(8)
            p.font.color.rgb = COLOR_TEXT_SECONDARY

    # Speaker Notes
    set_speaker_notes(slide, """
SLIDE 8: ASYMMETRIC CRYPTOGRAPHY (SPEAKER SCRIPT)

• CORE TALKING POINTS:
1. "In 1976, Whitfield Diffie and Martin Hellman introduced an idea that changed human history: Asymmetric or Public-Key Cryptography. Instead of a single key, every user has TWO keys."
2. "How does it work? Imagine Bob has a mailbox with an open mail slot. Anyone in the city can drop a letter inside through the open slot—that is Bob's Public Key. But only Bob has the physical key to unlock the mailbox door and read the letters—that is Bob's Private Key."
3. "Point out RSA and ECC:
   - RSA relies on prime factorization: Multiplying two 1000-digit prime numbers is instantaneous, but finding the original primes from their product takes millions of CPU hours.
   - Elliptic Curve Cryptography (ECC) is modern magic: It uses point arithmetic on algebraic curves. A 256-bit ECC key offers the exact same mathematical security as a massive 3072-bit RSA key, while consuming 90% less battery and memory. That is why your smartphone uses ECC!"
4. "Notice the comparison table on the right, and the bottom box: 'The Hybrid Cryptosystem'. In the real world, we don't choose between Symmetric and Asymmetric—we combine them! Asymmetric handles the key exchange at the start, and Symmetric encrypts the data packets."

• TRANSITION:
"Now, what happens if we need to verify that a file was NOT altered, or prove who authored it? For that, we need Hash Functions and Digital Signatures."
""")


# ==============================================================================
# SLIDE 9: HASH FUNCTIONS & DIGITAL SIGNATURES
# ==============================================================================
def build_slide_9(prs):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    apply_background(slide)
    add_header(
        slide,
        badge_text="MODULE 02: INTEGRITY & AUTHENTICITY",
        title_text="Hash Functions & Digital Signatures: Guaranteeing Trust",
        subtitle_text="Providing unforgeable data integrity, entity authentication, and legal non-repudiation",
        slide_num=9
    )

    card_w = Inches(5.7)
    card_h = Inches(5.15)
    start_y = Inches(1.65)

    # ---------------------------------------------------------
    # LEFT COLUMN: CRYPTOGRAPHIC HASH FUNCTIONS
    # ---------------------------------------------------------
    c_hash = add_card(slide, Inches(0.8), start_y, card_w, card_h, border_color=COLOR_CYAN_ACCENT)
    tb_h = slide.shapes.add_textbox(Inches(0.95), start_y + Inches(0.12), card_w - Inches(0.3), card_h - Inches(0.24))
    tf_h = tb_h.text_frame
    tf_h.word_wrap = True

    p_hh = tf_h.paragraphs[0]
    p_hh.text = "1. CRYPTOGRAPHIC HASH FUNCTIONS (DIGITAL FINGERPRINTS)"
    p_hh.font.name = FONT_HEADING
    p_hh.font.size = Pt(11)
    p_hh.font.bold = True
    p_hh.font.color.rgb = COLOR_CYAN_LIGHT

    p_hd = tf_h.add_paragraph()
    p_hd.space_before = Pt(3)
    p_hd.text = "Mathematical one-way mapping transforming arbitrary-length data into a fixed-length digest: h = H(M)."
    p_hd.font.name = FONT_BODY
    p_hd.font.size = Pt(8.8)
    p_hd.font.color.rgb = COLOR_TEXT_MUTED

    hash_props = [
        ("Deterministic:", "An identical input string or multi-gigabyte file will ALWAYS produce the exact same cryptographic hash output."),
        ("Pre-Image Resistance (One-Way):", "Given a hash digest h, it is computationally impossible to determine the original message M such that H(M) = h."),
        ("Small Avalanche Effect:", "Changing a single bit in the input alters over 50% of the output digest bits unpredictably, preventing statistical pattern analysis."),
        ("Collision Resistance:", "Computationally infeasible to find ANY two distinct inputs x and y such that H(x) == H(y). (Pigeonhole principle safeguard)."),
        ("Standards & Deprecation:", "SHA-256 and SHA-3 (Keccak sponge) are industry standards. MD5 (128-bit) and SHA-1 (160-bit) are cryptographically broken due to practical collision exploits.")
    ]
    for lbl, desc in hash_props:
        p = tf_h.add_paragraph()
        p.space_before = Pt(3)
        r1 = p.add_run()
        r1.text = f"• {lbl} "
        r1.font.bold = True
        r1.font.size = Pt(8.5)
        r1.font.color.rgb = COLOR_BLUE_ACCENT
        r2 = p.add_run()
        r2.text = desc
        r2.font.size = Pt(8.5)
        r2.font.color.rgb = COLOR_TEXT_SECONDARY

    # Hash visual box inside left card
    h_box = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.95), Inches(5.8), card_w - Inches(0.3), Inches(0.85))
    h_box.fill.solid()
    h_box.fill.fore_color.rgb = COLOR_INNER_BOX
    h_box.line.color.rgb = COLOR_CYAN_ACCENT
    tf_hb = h_box.text_frame
    tf_hb.word_wrap = True
    p = tf_hb.paragraphs[0]
    p.text = "SHA-256 HASH EXAMPLE (AVALANCHE DEMO):"
    p.font.name = FONT_MONO
    p.font.size = Pt(8)
    p.font.bold = True
    p.font.color.rgb = COLOR_CYAN_LIGHT
    p2 = tf_hb.add_paragraph()
    p2.text = 'H("Hello World")  -> a591a6d40bf420404a011733cfb7b190d62c65bf...\nH("Hello World.") -> 00f91b5c92be2a7c4e5a9609f7a974b62e49c716...'
    p2.font.name = FONT_MONO
    p2.font.size = Pt(7.5)
    p2.font.color.rgb = COLOR_TEXT_PRIMARY

    # ---------------------------------------------------------
    # RIGHT COLUMN: DIGITAL SIGNATURES & VERIFICATION FLOW
    # ---------------------------------------------------------
    rx = Inches(6.8)
    c_sig = add_card(slide, rx, start_y, card_w, card_h, border_color=COLOR_GREEN_ACCENT)
    tb_s = slide.shapes.add_textbox(rx + Inches(0.15), start_y + Inches(0.12), card_w - Inches(0.3), card_h - Inches(0.24))
    tf_s = tb_s.text_frame
    tf_s.word_wrap = True

    p_sh = tf_s.paragraphs[0]
    p_sh.text = "2. DIGITAL SIGNATURES: MATHEMATICAL PROOF OF TRUST"
    p_sh.font.name = FONT_HEADING
    p_sh.font.size = Pt(11)
    p_sh.font.bold = True
    p_sh.font.color.rgb = COLOR_GREEN_ACCENT

    p_sd = tf_s.add_paragraph()
    p_sd.space_before = Pt(3)
    p_sd.text = "Combining asymmetric encryption with hash functions to prove origin, authenticity, and non-repudiation."
    p_sd.font.name = FONT_BODY
    p_sd.font.size = Pt(8.8)
    p_sd.font.color.rgb = COLOR_TEXT_MUTED

    sig_steps = [
        ("Step 1 (Sender Signing):", "Sender computes hash digest h = H(Message). Sender encrypts digest h using their PRIVATE KEY: Signature = E_Kpriv(h). Signature is appended to message."),
        ("Step 2 (Transmission):", "Message and Signature travel together across the open internet to the recipient."),
        ("Step 3 (Verification):", "Recipient decrypts signature using Sender's PUBLIC KEY to retrieve original hash h. Recipient computes hash h' of received message. If h == h', signature is 100% VALID!"),
        ("The Triple Guarantee:", ""),
        ("  ✔ Integrity:", "If an adversary altered even 1 bit of the message, h != h', failing verification."),
        ("  ✔ Authenticity:", "Only the true sender possessing the secret private key could generate the signature."),
        ("  ✔ Non-Repudiation:", "Sender cannot legally or technically deny signing the transmission.")
    ]
    for lbl, desc in sig_steps:
        p = tf_s.add_paragraph()
        p.space_before = Pt(2)
        r1 = p.add_run()
        r1.text = f"{lbl} "
        r1.font.bold = True
        r1.font.size = Pt(8.5)
        r1.font.color.rgb = COLOR_GREEN_ACCENT if "✔" in lbl else COLOR_CYAN_LIGHT
        r2 = p.add_run()
        r2.text = desc
        r2.font.size = Pt(8.5)
        r2.font.color.rgb = COLOR_TEXT_SECONDARY

    # Bottom Right Photo embedded nicely in card
    add_photo_card(
        slide,
        'assets/slide9_hash.jpg',
        rx + Inches(0.15), Inches(5.35), card_w - Inches(0.3), Inches(1.3),
        caption="Figure 9.1: Cryptographic hash ledger & digital verification nodes.",
        border_color=COLOR_GREEN_ACCENT
    )

    # Speaker Notes
    set_speaker_notes(slide, """
SLIDE 9: HASH FUNCTIONS & DIGITAL SIGNATURES (SPEAKER SCRIPT)

• CORE TALKING POINTS:
1. "Students often confuse Encryption with Hashing. Let's make the distinction crystal clear:
   - Encryption is a two-way mathematical process: you encrypt with a key, and decrypt with a key.
   - Hashing is a ONE-WAY mathematical trapdoor: you can turn a 1,000-page book into a 256-bit hash, but you can NEVER reverse the hash to recreate the book."
2. "Look at the hash example on the left: Notice that adding just a single period at the end of 'Hello World' causes the entire hash digest to change completely! That is called the Avalanche Effect."
3. "Now look at the right side: Digital Signatures. How do they work?
   - In asymmetric encryption, you encrypt with the recipient's public key for privacy.
   - In digital signatures, we flip it! The sender encrypts the hash of the message with their OWN PRIVATE KEY.
   - Anyone on Earth can use the sender's public key to verify it.
   - Why encrypt the hash instead of the whole file? Efficiency! Computing an asymmetric cipher on a 10 GB software installer would take minutes; hashing it first and encrypting the 32-byte hash takes milliseconds."
4. "Digital signatures provide Non-Repudiation: In a court of law or banking audit, you cannot say 'someone else sent that transaction' because only your private key could have produced that signature."

• TRANSITION:
"Now let's see how these cryptographic primitives come together to secure the internet protocols we use every day."
""")


# ==============================================================================
# SLIDE 10: CRYPTOGRAPHIC PROTOCOLS IN ACTION
# ==============================================================================
def build_slide_10(prs):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    apply_background(slide)
    add_header(
        slide,
        badge_text="MODULE 02: APPLIED PROTOCOLS",
        title_text="Cryptographic Protocols in Action: Securing the Internet",
        subtitle_text="How layered cryptographic protocols protect real-world communications across untrusted backbones",
        slide_num=10
    )

    # 4 Structured Protocol Cards on Left (Width: 7.3)
    left_w = Inches(7.3)
    p_h = Inches(1.18)
    gap_y = Inches(0.12)
    start_y = Inches(1.65)

    protocols = [
        (
            "1. Transport Layer Security (TLS 1.3 / HTTPS)",
            COLOR_CYAN_LIGHT,
            "Port 443  •  Securing Web Commerce & APIs",
            "The protocol behind the browser padlock. Uses Ephemeral Elliptic Curve Diffie-Hellman (ECDHE) for instantaneous 1-RTT key exchange, AES-256-GCM for symmetric session encryption, and X.509 digital certificates for server authentication. Mandates Perfect Forward Secrecy (PFS)."
        ),
        (
            "2. Secure Shell Protocol (SSH-2)",
            COLOR_BLUE_ACCENT,
            "Port 22  •  Encrypted Remote System Administration",
            "Replaced legacy, unencrypted Telnet and rlogin. Employs asymmetric key pairs (ed25519 or RSA) for client authentication, establishing an encrypted tunnel using ChaCha20-Poly1305 or AES-CTR to protect terminal commands and automated SFTP file transfers."
        ),
        (
            "3. IPsec & Virtual Private Networks (VPNs)",
            COLOR_GREEN_ACCENT,
            "Layer 3 (Network Layer)  •  Enterprise Tunneling",
            "Encrypts all communication at the IP packet level. Uses Encapsulating Security Payload (ESP) to encrypt the entire original IP packet and Internet Key Exchange (IKEv2) to negotiate security associations, securely extending corporate intranets across public backbones."
        ),
        (
            "4. End-to-End Encryption (E2EE - Signal Protocol)",
            COLOR_PURPLE_ACCENT,
            "Application Layer  •  Messaging Security (Signal, WhatsApp)",
            "Employs the revolutionary Double Ratchet Algorithm: continuously deriving fresh, ephemeral cryptographic keys for every single message sent. Even if an adversary compromises a device's long-term key, prior message history remains mathematically uncrackable."
        )
    ]

    for idx, (p_title, p_col, p_sub, p_desc) in enumerate(protocols):
        py = start_y + idx * (p_h + gap_y)
        c = add_card(slide, Inches(0.8), py, left_w, p_h, border_color=COLOR_CARD_BORDER)
        
        # Color indicator strip
        strip = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0.8), py, Inches(0.08), p_h)
        strip.fill.solid()
        strip.fill.fore_color.rgb = p_col
        strip.line.fill.background()

        tb = slide.shapes.add_textbox(Inches(1.0), py + Inches(0.06), left_w - Inches(0.3), p_h - Inches(0.12))
        tf = tb.text_frame
        tf.word_wrap = True
        
        p = tf.paragraphs[0]
        r1 = p.add_run()
        r1.text = f"{p_title}  "
        r1.font.name = FONT_HEADING
        r1.font.size = Pt(10)
        r1.font.bold = True
        r1.font.color.rgb = p_col

        r2 = p.add_run()
        r2.text = f"[{p_sub}]"
        r2.font.name = FONT_MONO
        r2.font.size = Pt(8)
        r2.font.color.rgb = COLOR_TEXT_MUTED

        p_desc_para = tf.add_paragraph()
        p_desc_para.space_before = Pt(2)
        p_desc_para.text = p_desc
        p_desc_para.font.name = FONT_BODY
        p_desc_para.font.size = Pt(8.5)
        p_desc_para.font.color.rgb = COLOR_TEXT_SECONDARY

    # ---------------------------------------------------------
    # RIGHT COLUMN: NETWORK INFRASTRUCTURE PHOTO & OSI MAPPING
    # ---------------------------------------------------------
    rx = Inches(8.4)
    rw = Inches(4.133)
    
    add_photo_card(
        slide,
        'assets/slide10_protocols.jpg',
        rx, start_y, rw, Inches(3.2),
        caption="Figure 10.1: Encrypted data packets transiting fiber-optic backbones via TLS & IPsec.",
        border_color=COLOR_CYAN_ACCENT
    )

    osi_card = add_card(slide, rx, Inches(5.2), rw, Inches(1.6), bg_color=COLOR_INNER_BOX, border_color=COLOR_BLUE_ACCENT)
    tb_o = slide.shapes.add_textbox(rx + Inches(0.12), Inches(5.26), rw - Inches(0.24), Inches(1.48))
    tf_o = tb_o.text_frame
    tf_o.word_wrap = True
    p_oh = tf_o.paragraphs[0]
    p_oh.text = "OSI SECURITY STACK MAPPING"
    p_oh.font.name = FONT_HEADING
    p_oh.font.size = Pt(9.5)
    p_oh.font.bold = True
    p_oh.font.color.rgb = COLOR_CYAN_LIGHT

    osi_layers = [
        ("Application (L7):", "HTTPS, SSH, S/MIME, PGP, Signal E2EE"),
        ("Transport (L4):", "TLS 1.3, DTLS (for UDP / WebRTC)"),
        ("Network (L3):", "IPsec (AH / ESP), WireGuard"),
        ("Data Link (L2):", "MACsec (IEEE 802.1AE LAN encryption)")
    ]
    for olbl, odesc in osi_layers:
        p = tf_o.add_paragraph()
        p.space_before = Pt(2)
        r1 = p.add_run()
        r1.text = f"• {olbl} "
        r1.font.bold = True
        r1.font.size = Pt(8.2)
        r1.font.color.rgb = COLOR_BLUE_ACCENT
        r2 = p.add_run()
        r2.text = odesc
        r2.font.size = Pt(8.2)
        r2.font.color.rgb = COLOR_TEXT_PRIMARY

    # Speaker Notes
    set_speaker_notes(slide, """
SLIDE 10: CRYPTOGRAPHIC PROTOCOLS IN ACTION (SPEAKER SCRIPT)

• CORE TALKING POINTS:
1. "How do billions of people safely browse websites, transfer billions of dollars, and send private messages without hackers stealing their credentials? Through standardized cryptographic protocols."
2. "Let's review the top four protocols:
   - TLS 1.3: The heart of the web. Compared to TLS 1.2, version 1.3 removed obsolete insecure ciphers (like RC4 and CBC mode) and slashed the handshake time from two round trips to just one round trip. It enforces Perfect Forward Secrecy: even if a server's private certificate is stolen 5 years from now, recorded historical traffic cannot be decrypted!
   - SSH: What every engineer uses to connect to cloud servers. It completely eliminates passwords by matching local public-private SSH keys against ~/.ssh/authorized_keys.
   - IPsec: Encrypts network layer traffic. Unlike TLS, which only encrypts web browser packets, IPsec encrypts all traffic from an enterprise machine back to the corporate headquarters.
   - Signal Protocol: Used by Signal and WhatsApp. It utilizes the Double Ratchet Algorithm. Every single text message derives a brand-new ephemeral key. If someone steals your phone right now, they cannot decrypt earlier messages because those keys were already discarded."
3. "Notice the OSI Security Stack on the bottom right: Security can be implemented at Layer 7, Layer 4, Layer 3, or Layer 2. A defense-in-depth architecture secures multiple layers simultaneously."

• TRANSITION:
"Now, encryption and protocols are only half the battle. How do systems know WHO is requesting access? Let us examine Identity, Authentication, and Zero Trust."
""")

print("Part 2 (Slides 6 to 10) compiled successfully.")
