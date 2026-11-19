from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.enum.shapes import MSO_SHAPE
from build_presentation import (
    apply_background, add_header, add_card, add_photo_card, set_speaker_notes,
    COLOR_BG_DARK, COLOR_CARD_BG, COLOR_CARD_BORDER, COLOR_CARD_HIGHLIGHT, COLOR_INNER_BOX,
    COLOR_CYAN_ACCENT, COLOR_CYAN_LIGHT, COLOR_BLUE_ACCENT, COLOR_GREEN_ACCENT, COLOR_RED_ACCENT, COLOR_AMBER_ACCENT, COLOR_PURPLE_ACCENT,
    COLOR_TEXT_PRIMARY, COLOR_TEXT_SECONDARY, COLOR_TEXT_MUTED,
    FONT_HEADING, FONT_BODY, FONT_MONO
)

# ==============================================================================
# SLIDE 1: TITLE SLIDE
# ==============================================================================
def build_slide_1(prs):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    apply_background(slide)

    # Decorative Cyan Accent Bar across top
    top_bar = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0.8), Inches(0.8), Inches(11.733), Inches(0.04))
    top_bar.fill.solid()
    top_bar.fill.fore_color.rgb = COLOR_CYAN_ACCENT
    top_bar.line.fill.background()

    # Seminar Department Tag
    dept_badge = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.8), Inches(1.05), Inches(5.8), Inches(0.35))
    dept_badge.fill.solid()
    dept_badge.fill.fore_color.rgb = COLOR_INNER_BOX
    dept_badge.line.color.rgb = COLOR_CYAN_ACCENT
    dept_badge.line.width = Pt(1.2)
    tf_d = dept_badge.text_frame
    tf_d.margin_left = Inches(0.12)
    p_d = tf_d.paragraphs[0]
    p_d.text = "DEPARTMENT OF COMPUTER SCIENCE & ENGINEERING  •  TECHNICAL SEMINAR"
    p_d.font.name = FONT_HEADING
    p_d.font.size = Pt(8.5)
    p_d.font.bold = True
    p_d.font.color.rgb = COLOR_CYAN_LIGHT

    # Main Title
    title_box = slide.shapes.add_textbox(Inches(0.8), Inches(1.55), Inches(7.5), Inches(1.4))
    tf_t = title_box.text_frame
    tf_t.word_wrap = True
    tf_t.margin_left = tf_t.margin_right = tf_t.margin_top = tf_t.margin_bottom = 0
    p_t = tf_t.paragraphs[0]
    p_t.text = "Cybersecurity and Cryptography Methods"
    p_t.font.name = FONT_HEADING
    p_t.font.size = Pt(30)
    p_t.font.bold = True
    p_t.font.color.rgb = COLOR_TEXT_PRIMARY

    # Subtitle
    sub_box = slide.shapes.add_textbox(Inches(0.8), Inches(2.85), Inches(7.5), Inches(0.55))
    tf_s = sub_box.text_frame
    tf_s.word_wrap = True
    tf_s.margin_left = tf_s.margin_right = tf_s.margin_top = tf_s.margin_bottom = 0
    p_s = tf_s.paragraphs[0]
    p_s.text = "Protecting Digital Information, Critical Systems, and Modern Networks in an Adversarial World"
    p_s.font.name = FONT_BODY
    p_s.font.size = Pt(13)
    p_s.font.color.rgb = COLOR_BLUE_ACCENT

    # 3 Pillar Badges
    pills = [
        ("01. THREAT ANALYSIS & CIA", COLOR_RED_ACCENT),
        ("02. MODERN CIPHERS & PROTOCOLS", COLOR_CYAN_ACCENT),
        ("03. ZERO TRUST & POST-QUANTUM", COLOR_GREEN_ACCENT)
    ]
    for i, (ptext, pcol) in enumerate(pills):
        pill_shape = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.8 + i * 2.5), Inches(3.55), Inches(2.38), Inches(0.35))
        pill_shape.fill.solid()
        pill_shape.fill.fore_color.rgb = COLOR_INNER_BOX
        pill_shape.line.color.rgb = pcol
        pill_shape.line.width = Pt(1)
        tf_p = pill_shape.text_frame
        p_p = tf_p.paragraphs[0]
        p_p.text = ptext
        p_p.alignment = PP_ALIGN.CENTER
        p_p.font.name = FONT_MONO
        p_p.font.size = Pt(8)
        p_p.font.bold = True
        p_p.font.color.rgb = pcol

    # Presenter Information Card (Editable placeholders)
    pres_card = add_card(slide, Inches(0.8), Inches(4.15), Inches(7.4), Inches(2.55), bg_color=COLOR_CARD_BG, border_color=COLOR_CARD_BORDER)
    pres_box = slide.shapes.add_textbox(Inches(1.05), Inches(4.3), Inches(6.9), Inches(2.25))
    tf_pr = pres_box.text_frame
    tf_pr.word_wrap = True
    tf_pr.margin_left = tf_pr.margin_right = tf_pr.margin_top = tf_pr.margin_bottom = 0

    p_h = tf_pr.paragraphs[0]
    p_h.text = "PRESENTER & ACADEMIC CREDENTIALS"
    p_h.font.name = FONT_HEADING
    p_h.font.size = Pt(10)
    p_h.font.bold = True
    p_h.font.color.rgb = COLOR_CYAN_LIGHT

    credentials = [
        ("Student Name:", "[Enter Student Name Here]"),
        ("Roll No. / USN:", "[Enter University Seat / Roll Number]"),
        ("Program & Branch:", "Bachelor of Engineering (B.E.) — Computer Science & Engineering"),
        ("Faculty Guide / Mentor:", "[Enter Guide Name, Designation, Dept. of CSE]"),
        ("Institution:", "[Enter Engineering College / University Name]"),
        ("Academic Year:", "2025 – 2026  |  Final Year Technical Seminar")
    ]
    for lbl, val in credentials:
        p = tf_pr.add_paragraph()
        p.space_before = Pt(3)
        run1 = p.add_run()
        run1.text = f"{lbl}  "
        run1.font.name = FONT_BODY
        run1.font.bold = True
        run1.font.size = Pt(9.5)
        run1.font.color.rgb = COLOR_TEXT_MUTED

        run2 = p.add_run()
        run2.text = val
        run2.font.name = FONT_BODY
        run2.font.bold = False
        run2.font.size = Pt(9.5)
        run2.font.color.rgb = COLOR_TEXT_PRIMARY

    # Right Hero Image & Caption
    add_photo_card(
        slide,
        'assets/slide1_hero.jpg',
        Inches(8.5), Inches(1.55), Inches(4.033), Inches(4.75),
        caption="Figure 1.1: Multi-vector cyber intelligence and cryptographic operations center.",
        border_color=COLOR_CYAN_ACCENT
    )

    # Footer line
    foot_box = slide.shapes.add_textbox(Inches(0.8), Inches(6.95), Inches(11.733), Inches(0.25))
    p_f = foot_box.text_frame.paragraphs[0]
    p_f.text = "Cybersecurity and Cryptography Methods  •  Department of Computer Science & Engineering"
    p_f.font.name = FONT_BODY
    p_f.font.size = Pt(8.5)
    p_f.font.color.rgb = COLOR_TEXT_MUTED

    # Speaker Notes
    set_speaker_notes(slide, """
SLIDE 1: TITLE SLIDE (SPEAKER SCRIPT)

• GREETING & HOOK:
"Respected faculty members, evaluators, and my fellow engineering peers, good morning. Today, I am presenting my technical seminar on 'Cybersecurity and Cryptography Methods: Protecting Digital Information in the Modern World'."

• MOTIVATION & CONTEXT:
"In today's hyper-connected computing landscape, digital data has become the world's most valuable asset. Whether it is financial transactions, healthcare telemetry, sovereign defense systems, or enterprise cloud workloads—our entire civilization runs on software and distributed networks. However, this hyper-connectivity has created an unprecedented attack surface."

• PRESENTATION ROADMAP:
"Throughout this 15-slide seminar, we will bridge two complementary pillars of computer science:
1. Cybersecurity: The holistic architectural policies, threat detection systems, and defensive frameworks that safeguard assets.
2. Cryptography: The deep mathematical bedrock of ciphers, hashing, and key protocols that mathematically guarantee confidentiality, integrity, and authenticity.
We will examine symmetric and asymmetric algorithms, cryptographic network protocols, identity mechanisms, real-world breach post-mortems, and conclude with the frontier challenges of Quantum computing and Artificial Intelligence."

• TRANSITION:
"Let us begin by defining the exact scope and pressing imperatives of modern cybersecurity."
""")


# ==============================================================================
# SLIDE 2: INTRODUCTION TO CYBERSECURITY
# ==============================================================================
def build_slide_2(prs):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    apply_background(slide)
    add_header(
        slide,
        badge_text="MODULE 01: FOUNDATIONS",
        title_text="Introduction to Cybersecurity: Scope & Imperatives",
        subtitle_text="Defending interconnected computational systems, protocols, and data from unauthorized exploitation",
        slide_num=2
    )

    # Left Section: 3 Structured Cards
    card_w = Inches(7.3)
    
    # Card 1: What is Cybersecurity?
    c1 = add_card(slide, Inches(0.8), Inches(1.65), card_w, Inches(1.55))
    tb1 = slide.shapes.add_textbox(Inches(0.95), Inches(1.72), Inches(7.0), Inches(1.4))
    tf1 = tb1.text_frame
    tf1.word_wrap = True
    p1 = tf1.paragraphs[0]
    p1.text = "1. Definitional Scope of Cybersecurity"
    p1.font.name = FONT_HEADING
    p1.font.size = Pt(11.5)
    p1.font.bold = True
    p1.font.color.rgb = COLOR_CYAN_LIGHT

    bullets1 = [
        "Systemic Discipline: The practice of protecting computers, servers, mobile devices, electronic systems, networks, and data from malicious attacks, unauthorized access, and operational disruptions.",
        "Engineering Philosophy: Security is not an isolated product or plugin; it is a holistic architectural lifecycle spanning people, processes, physical infrastructure, and mathematical primitives."
    ]
    for b in bullets1:
        p = tf1.add_paragraph()
        p.space_before = Pt(3)
        p.text = f"•  {b}"
        p.font.name = FONT_BODY
        p.font.size = Pt(9.5)
        p.font.color.rgb = COLOR_TEXT_SECONDARY

    # Card 2: Why Digital Protection is Critical
    c2 = add_card(slide, Inches(0.8), Inches(3.3), card_w, Inches(1.6))
    tb2 = slide.shapes.add_textbox(Inches(0.95), Inches(3.37), Inches(7.0), Inches(1.45))
    tf2 = tb2.text_frame
    tf2.word_wrap = True
    p2 = tf2.paragraphs[0]
    p2.text = "2. Why Digital Information Demands Rigorous Protection"
    p2.font.name = FONT_HEADING
    p2.font.size = Pt(11.5)
    p2.font.bold = True
    p2.font.color.rgb = COLOR_BLUE_ACCENT

    bullets2 = [
        "Astronomical Financial Fallout: Global cybercrime damages are projected to exceed $10.5 Trillion annually, driven by organized ransomware cartels, intellectual property espionage, and extortion.",
        "Critical Infrastructure Interdependence: Modern attacks target power grids, water purification, healthcare systems, and air traffic control where breaches directly jeopardize human life.",
        "Regulatory & Legal Mandates: Stringent data compliance standards (GDPR, HIPAA, ISO/IEC 27001, DPDPA) carry severe statutory penalties for compromised consumer privacy."
    ]
    for b in bullets2:
        p = tf2.add_paragraph()
        p.space_before = Pt(3)
        p.text = f"•  {b}"
        p.font.name = FONT_BODY
        p.font.size = Pt(9.5)
        p.font.color.rgb = COLOR_TEXT_SECONDARY

    # Card 3: The 4 Protective Domains
    c3 = add_card(slide, Inches(0.8), Inches(5.0), card_w, Inches(1.85))
    tb3 = slide.shapes.add_textbox(Inches(0.95), Inches(5.07), Inches(7.0), Inches(1.7))
    tf3 = tb3.text_frame
    tf3.word_wrap = True
    p3 = tf3.paragraphs[0]
    p3.text = "3. The Four Core Computational Frontiers"
    p3.font.name = FONT_HEADING
    p3.font.size = Pt(11.5)
    p3.font.bold = True
    p3.font.color.rgb = COLOR_GREEN_ACCENT

    frontiers = [
        ("Compute & Endpoints:", "Laptops, workstations, mobile devices, IoT micro-controllers, and bare-metal enterprise servers."),
        ("Networks & Routing:", "Perimeter firewalls, software-defined networks (SDN), encrypted packet switching, and DNS."),
        ("Application Logic:", "Web applications, microservices, REST APIs, and database engines resistant to code injection."),
        ("Data States:", "Safeguarding data 'at rest' (storage), 'in transit' (networks), and 'in use' (RAM / secure enclaves).")
    ]
    for f_title, f_desc in frontiers:
        p = tf3.add_paragraph()
        p.space_before = Pt(2)
        r1 = p.add_run()
        r1.text = f"•  {f_title} "
        r1.font.bold = True
        r1.font.size = Pt(9.2)
        r1.font.color.rgb = COLOR_CYAN_LIGHT
        r2 = p.add_run()
        r2.text = f_desc
        r2.font.size = Pt(9.2)
        r2.font.color.rgb = COLOR_TEXT_SECONDARY

    # Right Column: Photo & Caption Card
    add_photo_card(
        slide,
        'assets/slide2_soc.jpg',
        Inches(8.4), Inches(1.65), Inches(4.133), Inches(3.2),
        caption="Figure 2.1: 24/7 Security Operations Center (SOC) monitoring enterprise telemetry.",
        border_color=COLOR_CYAN_ACCENT
    )

    # Everyday Examples Callout Card
    ex_card = add_card(slide, Inches(8.4), Inches(5.3), Inches(4.133), Inches(1.55), bg_color=COLOR_CARD_HIGHLIGHT, border_color=COLOR_CYAN_ACCENT)
    ex_tb = slide.shapes.add_textbox(Inches(8.55), Inches(5.38), Inches(3.833), Inches(1.4))
    tf_ex = ex_tb.text_frame
    tf_ex.word_wrap = True
    p_ex_h = tf_ex.paragraphs[0]
    p_ex_h.text = "EVERYDAY CYBERSECURITY EXAMPLES"
    p_ex_h.font.name = FONT_HEADING
    p_ex_h.font.size = Pt(10)
    p_ex_h.font.bold = True
    p_ex_h.font.color.rgb = COLOR_CYAN_LIGHT

    examples = [
        "Banking OTPs & UPI: Cryptographic tokens validating financial transactions.",
        "HTTPS Browser Padlocks: TLS handshakes securing web commerce.",
        "Biometric Phone Unlocks: Hardware-backed secure enclaves verifying fingerprints.",
        "Automatic OS Updates: Patching kernel vulnerabilities before zero-day weaponization."
    ]
    for ex in examples:
        p = tf_ex.add_paragraph()
        p.space_before = Pt(2)
        p.text = f"✔ {ex}"
        p.font.name = FONT_BODY
        p.font.size = Pt(8.5)
        p.font.color.rgb = COLOR_TEXT_PRIMARY

    # Speaker Notes
    set_speaker_notes(slide, """
SLIDE 2: INTRODUCTION TO CYBERSECURITY (SPEAKER SCRIPT)

• CORE TALKING POINTS:
1. "To understand cybersecurity, we must realize that digital information is no longer confined to isolated mainframe rooms. It travels across global fiber cables, cellular towers, and edge cloud nodes."
2. "Why is cybersecurity so critical in 2026? Look at the numbers: cybercrime damages represent the third largest economy in the world after the US and China. Moreover, attacks are no longer just teenage pranksters defacing websites—they are state-sponsored advanced persistent threat (APT) actors and multi-million-dollar ransomware cartels targeting national infrastructure, hospitals, and supply chains."
3. "Notice the four frontiers: Endpoints, Networks, Applications, and Data. Notice especially the three states of data: data at rest on solid-state drives, data in transit traversing the internet, and data in use in RAM and CPU registers. Each state requires distinct defensive controls."
4. "Every single student in this room experiences cybersecurity dozens of times every day—when you receive a banking OTP, when your WhatsApp chats display 'end-to-end encrypted', or when your phone checks biometric hashes inside its secure enclave."

• AUDIENCE INTERACTION / EVALUATOR NOTE:
"If asked why perimeter firewalls alone are no longer sufficient, explain that the perimeter dissolved with remote work, mobile devices, and multi-cloud architectures, requiring defense-in-depth."
""")


# ==============================================================================
# SLIDE 3: THE CIA TRIAD
# ==============================================================================
def build_slide_3(prs):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    apply_background(slide)
    add_header(
        slide,
        badge_text="MODULE 01: CORE ARCHITECTURE",
        title_text="The CIA Triad: Foundational Information Security Paradigm",
        subtitle_text="The three interdependent principles governing every security policy, cryptographic control, and architectural design",
        slide_num=3
    )

    col_w = Inches(3.75)
    gap = Inches(0.24)
    top_pos = Inches(1.65)
    col_h = Inches(3.45)

    # Column 1: Confidentiality
    c1 = add_card(slide, Inches(0.8), top_pos, col_w, col_h, border_color=COLOR_CYAN_ACCENT, border_width=1.5)
    tb1 = slide.shapes.add_textbox(Inches(0.95), top_pos + Inches(0.1), col_w - Inches(0.3), col_h - Inches(0.2))
    tf1 = tb1.text_frame
    tf1.word_wrap = True
    p1 = tf1.paragraphs[0]
    p1.text = "CONFIDENTIALITY (C)"
    p1.font.name = FONT_HEADING
    p1.font.size = Pt(12)
    p1.font.bold = True
    p1.font.color.rgb = COLOR_CYAN_LIGHT

    c_points = [
        ("Core Mandate:", "Ensuring sensitive information is accessible ONLY to authenticated and authorized principals; zero unauthorized exposure."),
        ("Cryptographic Controls:", "AES-256 symmetric ciphers, RSA/ECC public key encryption, TLS 1.3 tunnels, database column encryption."),
        ("Architectural Defenses:", "Role-Based Access Control (RBAC), multi-factor authentication (MFA), data classification, tokenization."),
        ("Violation Vectors:", "Eavesdropping, packet sniffing, credential harvesting, unauthorized database dumps, insider exfiltration.")
    ]
    for lbl, desc in c_points:
        p = tf1.add_paragraph()
        p.space_before = Pt(4)
        r1 = p.add_run()
        r1.text = f"• {lbl} "
        r1.font.bold = True
        r1.font.size = Pt(8.8)
        r1.font.color.rgb = COLOR_CYAN_ACCENT
        r2 = p.add_run()
        r2.text = desc
        r2.font.size = Pt(8.8)
        r2.font.color.rgb = COLOR_TEXT_SECONDARY

    # Column 2: Integrity
    c2 = add_card(slide, Inches(0.8) + col_w + gap, top_pos, col_w, col_h, border_color=COLOR_GREEN_ACCENT, border_width=1.5)
    tb2 = slide.shapes.add_textbox(Inches(0.8) + col_w + gap + Inches(0.15), top_pos + Inches(0.1), col_w - Inches(0.3), col_h - Inches(0.2))
    tf2 = tb2.text_frame
    tf2.word_wrap = True
    p2 = tf2.paragraphs[0]
    p2.text = "INTEGRITY (I)"
    p2.font.name = FONT_HEADING
    p2.font.size = Pt(12)
    p2.font.bold = True
    p2.font.color.rgb = COLOR_GREEN_ACCENT

    i_points = [
        ("Core Mandate:", "Guaranteeing that information remains accurate, complete, unaltered, and protected against unauthorized modification or deletion."),
        ("Cryptographic Controls:", "Cryptographic hash functions (SHA-256, SHA-3), Hash-based Message Authentication Codes (HMAC), digital signatures."),
        ("Architectural Defenses:", "Immutable audit logs, version control checkouts, cryptographic checksums, file integrity monitoring (FIM)."),
        ("Violation Vectors:", "Man-in-the-Middle (MITM) packet tampering, SQL injection data corruption, unauthorized code injection, replay attacks.")
    ]
    for lbl, desc in i_points:
        p = tf2.add_paragraph()
        p.space_before = Pt(4)
        r1 = p.add_run()
        r1.text = f"• {lbl} "
        r1.font.bold = True
        r1.font.size = Pt(8.8)
        r1.font.color.rgb = COLOR_GREEN_ACCENT
        r2 = p.add_run()
        r2.text = desc
        r2.font.size = Pt(8.8)
        r2.font.color.rgb = COLOR_TEXT_SECONDARY

    # Column 3: Availability
    c3 = add_card(slide, Inches(0.8) + (col_w + gap) * 2, top_pos, col_w, col_h, border_color=COLOR_AMBER_ACCENT, border_width=1.5)
    tb3 = slide.shapes.add_textbox(Inches(0.8) + (col_w + gap) * 2 + Inches(0.15), top_pos + Inches(0.1), col_w - Inches(0.3), col_h - Inches(0.2))
    tf3 = tb3.text_frame
    tf3.word_wrap = True
    p3 = tf3.paragraphs[0]
    p3.text = "AVAILABILITY (A)"
    p3.font.name = FONT_HEADING
    p3.font.size = Pt(12)
    p3.font.bold = True
    p3.font.color.rgb = COLOR_AMBER_ACCENT

    a_points = [
        ("Core Mandate:", "Ensuring authorized systems and verified users possess timely, reliable, and uninterrupted access to critical information and services."),
        ("Cryptographic Controls:", "Efficient lightweight ciphers preventing computational exhaustion, cryptographically validated failover handshakes."),
        ("Architectural Defenses:", "Multi-region redundant clustering, automated load balancers, Anycast DDoS mitigation, RAID storage arrays."),
        ("Violation Vectors:", "Volumetric Distributed Denial of Service (DDoS), ransomware encryption of operational databases, physical outages.")
    ]
    for lbl, desc in a_points:
        p = tf3.add_paragraph()
        p.space_before = Pt(4)
        r1 = p.add_run()
        r1.text = f"• {lbl} "
        r1.font.bold = True
        r1.font.size = Pt(8.8)
        r1.font.color.rgb = COLOR_AMBER_ACCENT
        r2 = p.add_run()
        r2.text = desc
        r2.font.size = Pt(8.8)
        r2.font.color.rgb = COLOR_TEXT_SECONDARY

    # Bottom Row: Left Photo + Right Engineering Balance Callout
    add_photo_card(
        slide,
        'assets/slide3_cia.jpg',
        Inches(0.8), Inches(5.25), Inches(4.2), Inches(1.6),
        caption="Figure 3.1: Resilient high-availability data center node maintaining operational uptime.",
        border_color=COLOR_BLUE_ACCENT
    )

    bal_card = add_card(slide, Inches(5.2), Inches(5.25), Inches(7.333), Inches(1.6), bg_color=COLOR_CARD_HIGHLIGHT, border_color=COLOR_CYAN_ACCENT)
    tb_bal = slide.shapes.add_textbox(Inches(5.35), Inches(5.33), Inches(7.0), Inches(1.45))
    tf_bal = tb_bal.text_frame
    tf_bal.word_wrap = True
    p_bh = tf_bal.paragraphs[0]
    p_bh.text = "THE SECURITY ENGINEER'S BALANCING DILEMMA"
    p_bh.font.name = FONT_HEADING
    p_bh.font.size = Pt(10)
    p_bh.font.bold = True
    p_bh.font.color.rgb = COLOR_CYAN_LIGHT

    bal_text = [
        "Trade-offs in System Design: Maximizing one pillar often strains another. For example, enforcing extreme cryptographic authentication layers to maximize Confidentiality can introduce significant latency, harming Availability.",
        "Context-Driven Prioritization: A military defense database prioritizes Confidentiality above all else. Conversely, an Intensive Care Unit (ICU) medical monitoring system prioritizes Availability and Integrity—a delayed record can prove fatal.",
        "The Parkerian Hexad Extension: Modern security expands the CIA Triad to include Authenticity, Possession (Control), and Utility."
    ]
    for bt in bal_text:
        p = tf_bal.add_paragraph()
        p.space_before = Pt(2)
        p.text = f"• {bt}"
        p.font.name = FONT_BODY
        p.font.size = Pt(8.5)
        p.font.color.rgb = COLOR_TEXT_SECONDARY

    # Speaker Notes
    set_speaker_notes(slide, """
SLIDE 3: THE CIA TRIAD (SPEAKER SCRIPT)

• CORE TALKING POINTS:
1. "The CIA Triad is the bedrock of information security. Every security policy, firewall rule, encryption cipher, or access control list ever created was designed to protect one or more of these three pillars: Confidentiality, Integrity, or Availability."
2. "Let's walk through them carefully:
   - Confidentiality means privacy: only authorized eyes can view the data. If a hacker intercepts an encrypted email and cannot decrypt it, Confidentiality holds.
   - Integrity means authenticity and accuracy: the data has not been modified in transit or storage. If an attacker tampers with a financial wire transfer of $1,000 to say $100,000, that is an Integrity violation. We use SHA-256 and digital signatures to detect this.
   - Availability means accessibility: the system is operational when needed. If a hospital's patient records are locked by ransomware or paralyzed by a DDoS flood, that is an Availability failure."
3. "Point out the bottom right card: In computer science, security is an engineering trade-off. You can make a computer 100% confidential by disconnecting it from the internet, locking it in a lead vault, and encasing it in concrete—but its Availability drops to zero! As software engineers, our job is balancing these three pillars based on the specific threat model."

• ANTICIPATED SEMINAR QUESTION:
"Can you give an example where Integrity is more important than Confidentiality?"
Answer: "Yes! In a public public-health advisory or stock exchange ticker, the information is completely public (zero confidentiality needed), but its integrity must be 100% guaranteed so nobody manipulates price feeds or emergency notices."
""")


# ==============================================================================
# SLIDE 4: COMMON CYBERSECURITY THREATS
# ==============================================================================
def build_slide_4(prs):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    apply_background(slide)
    add_header(
        slide,
        badge_text="MODULE 01: THREAT INTELLIGENCE",
        title_text="Modern Cyber Threat Taxonomy: How Systems Fall",
        subtitle_text="Deconstructing primary adversarial vectors to engineer resilient, defense-oriented software systems",
        slide_num=4
    )

    # 6 Threat Cards in a 2-Column Matrix on Left (Width: 7.2)
    t_width = Inches(3.5)
    t_height = Inches(1.65)
    gap_x = Inches(0.2)
    gap_y = Inches(0.12)
    start_x = Inches(0.8)
    start_y = Inches(1.65)

    threats = [
        (
            "01. Malware & Rootkits",
            COLOR_RED_ACCENT,
            "Malicious binaries, spyware, and kernel rootkits executing unauthorized code, bypassing host defenses, establishing persistence, and opening Command & Control (C2) communication channels."
        ),
        (
            "02. Ransomware Extortion",
            COLOR_RED_ACCENT,
            "Malicious automated cryptographic locking of organizational file systems (using AES/RSA) followed by cryptocurrency extortion demands (e.g., LockBit, BlackCat) and double-extortion data leak threats."
        ),
        (
            "03. Phishing & Deception",
            COLOR_AMBER_ACCENT,
            "Adversarial email, SMS (smishing), and voice lures exploiting cognitive trust to steal session tokens, corporate credentials, or trigger malicious macro downloads. Accounts for over 80% of initial breaches."
        ),
        (
            "04. Password & Credential Attacks",
            COLOR_AMBER_ACCENT,
            "Automated brute-force attacks, dictionary sprays, rainbow-table cracking of unsalted hashes, and pass-the-hash lateral movement exploiting weak credential reuse across multiple systems."
        ),
        (
            "05. Social Engineering",
            COLOR_CYAN_LIGHT,
            "Direct psychological manipulation of human operators through pretexting, tailgating into data centers, and executive impersonation (Business Email Compromise - BEC) to bypass technical firewalls."
        ),
        (
            "06. DoS & Distributed DDoS",
            COLOR_CYAN_LIGHT,
            "Flooding network pipes and server connection pools (SYN floods, DNS/NTP amplification) to exhaust memory and CPU bandwidth, causing complete service starvation for legitimate users."
        )
    ]

    for idx, (ttitle, tcol, tdesc) in enumerate(threats):
        row = idx // 2
        col = idx % 2
        tx = start_x + col * (t_width + gap_x)
        ty = start_y + row * (t_height + gap_y)

        c = add_card(slide, tx, ty, t_width, t_height, border_color=COLOR_CARD_BORDER)
        tb = slide.shapes.add_textbox(tx + Inches(0.12), ty + Inches(0.08), t_width - Inches(0.24), t_height - Inches(0.15))
        tf = tb.text_frame
        tf.word_wrap = True
        p_th = tf.paragraphs[0]
        p_th.text = ttitle
        p_th.font.name = FONT_HEADING
        p_th.font.size = Pt(10.5)
        p_th.font.bold = True
        p_th.font.color.rgb = tcol

        p_td = tf.add_paragraph()
        p_td.space_before = Pt(2)
        p_td.text = tdesc
        p_td.font.name = FONT_BODY
        p_td.font.size = Pt(8.5)
        p_td.font.color.rgb = COLOR_TEXT_SECONDARY

    # Right Column: 2 Visuals (Phishing & Ransomware) + Countermeasure Badge
    rx = Inches(8.3)
    rw = Inches(4.233)
    
    # Photo 1: Phishing
    add_photo_card(
        slide,
        'assets/slide4_phishing.jpg',
        rx, Inches(1.65), rw, Inches(1.7),
        caption="Figure 4.1: Social deception lure harvesting enterprise credentials.",
        border_color=COLOR_AMBER_ACCENT
    )

    # Photo 2: Ransomware
    add_photo_card(
        slide,
        'assets/slide4_ransomware.jpg',
        rx, Inches(3.8), rw, Inches(1.6),
        caption="Figure 4.2: Adversarial ransomware terminal demanding decryption extortion.",
        border_color=COLOR_RED_ACCENT
    )

    # Countermeasure Card at Bottom Right
    count_card = add_card(slide, rx, Inches(5.8), rw, Inches(1.05), bg_color=COLOR_CARD_HIGHLIGHT, border_color=COLOR_CYAN_ACCENT)
    tb_c = slide.shapes.add_textbox(rx + Inches(0.12), Inches(5.85), rw - Inches(0.24), Inches(0.95))
    tf_c = tb_c.text_frame
    tf_c.word_wrap = True
    p_ch = tf_c.paragraphs[0]
    p_ch.text = "ENGINEERING COUNTERMEASURE MATRIX"
    p_ch.font.name = FONT_HEADING
    p_ch.font.size = Pt(9.5)
    p_ch.font.bold = True
    p_ch.font.color.rgb = COLOR_CYAN_LIGHT

    p_cd = tf_c.add_paragraph()
    p_cd.space_before = Pt(2)
    p_cd.text = "Mitigation Mandate: Zero-Trust verification, hardware MFA (FIDO2), endpoint behavioral heuristics (EDR), and immutable offline air-gapped backups."
    p_cd.font.name = FONT_BODY
    p_cd.font.size = Pt(8.5)
    p_cd.font.color.rgb = COLOR_TEXT_PRIMARY

    # Speaker Notes
    set_speaker_notes(slide, """
SLIDE 4: COMMON CYBERSECURITY THREATS (SPEAKER SCRIPT)

• CORE TALKING POINTS:
1. "To build strong defenses, computer science engineers must understand the mind and mechanics of an adversary. On this slide, we classify the six most destructive attack vectors."
2. "Notice how attackers chain these threats together into a Cyber Kill Chain:
   - It usually begins with Threat #3 (Phishing) or Threat #5 (Social Engineering). An employee clicks a malicious invoice link.
   - That drops Threat #1 (Malware/Trojan) onto the machine.
   - The malware executes Threat #4 (Credential harvesting), dumping password hashes from memory.
   - The attacker moves laterally across the enterprise network until they reach domain controller servers.
   - Finally, they detonate Threat #2 (Ransomware), encrypting millions of customer records and demanding millions in ransom."
3. "Notice Threat #6: Distributed Denial of Service. Unlike malware, DDoS doesn't seek to steal data; it seeks to destroy Availability by overwhelming server bandwidth."
4. "Point to the bottom right countermeasure box: Note that traditional antivirus software based on signature scanning is obsolete against polymorphic malware. Modern security requires Endpoint Detection and Response (EDR) utilizing behavioral heuristics and hardware MFA."

• KEY SEMINAR TAKEAWAY:
"More than 85% of successful organizational breaches exploit human vulnerability rather than zero-day flaws in mathematical encryption algorithms."
""")


# ==============================================================================
# SLIDE 5: TYPES OF CYBERSECURITY
# ==============================================================================
def build_slide_5(prs):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    apply_background(slide)
    add_header(
        slide,
        badge_text="MODULE 01: DOMAIN TAXONOMY",
        title_text="Five Core Disciplines of Enterprise Cybersecurity",
        subtitle_text="Comprehensive defense across the entire technology stack from physical silicon to distributed cloud workloads",
        slide_num=5
    )

    # 5 Structured Domain Cards on Left
    domains = [
        (
            "1. Network Security",
            COLOR_CYAN_LIGHT,
            "Perimeter & Traffic Defense",
            "Next-Generation Firewalls (NGFW), stateful packet inspection, Intrusion Detection and Prevention Systems (IDS/IPS), Network Segmentation via VLANs, and encrypted tunneling protocols (IPsec, WireGuard) safeguarding transmission pipelines."
        ),
        (
            "2. Application Security",
            COLOR_BLUE_ACCENT,
            "Secure Software Development Lifecycle",
            "Integrating security into the code pipeline: Static and Dynamic Application Security Testing (SAST/DAST), Software Composition Analysis (SCA), API security gateways, and rigorous input sanitization protecting against OWASP Top 10 flaws."
        ),
        (
            "3. Endpoint Security",
            COLOR_GREEN_ACCENT,
            "Device & Host Hardening",
            "Next-Gen Antivirus (NGAV), Endpoint Detection and Response (EDR) sensors monitoring kernel behavior, automated device disk encryption (BitLocker, FileVault), and Mobile Device Management (MDM) enforcing zero-trust posture."
        ),
        (
            "4. Cloud Security",
            COLOR_PURPLE_ACCENT,
            "Virtual Workloads & Shared Responsibility",
            "Cloud Security Posture Management (CSPM), Identity and Access Management (IAM) role federation, container image scanning (Docker, Kubernetes), and safeguarding microservices in multi-tenant cloud environments (AWS, Azure, GCP)."
        ),
        (
            "5. Data Security",
            COLOR_AMBER_ACCENT,
            "Cryptographic Information Governance",
            "Data Loss Prevention (DLP) engines, database encryption at-rest (AES-256), cryptographic tokenization, granular role-based access control (RBAC), and physical Hardware Security Modules (HSMs) isolating master cryptographic keys."
        )
    ]

    card_w = Inches(7.3)
    card_h = Inches(0.92)
    start_y = Inches(1.65)
    gap_y = Inches(0.12)

    for idx, (dtitle, dcol, dsub, ddesc) in enumerate(domains):
        cy = start_y + idx * (card_h + gap_y)
        c = add_card(slide, Inches(0.8), cy, card_w, card_h, border_color=COLOR_CARD_BORDER)
        
        # Left color indicator strip
        strip = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0.8), cy, Inches(0.08), card_h)
        strip.fill.solid()
        strip.fill.fore_color.rgb = dcol
        strip.line.fill.background()

        tb = slide.shapes.add_textbox(Inches(1.0), cy + Inches(0.06), card_w - Inches(0.3), card_h - Inches(0.12))
        tf = tb.text_frame
        tf.word_wrap = True
        p_dh = tf.paragraphs[0]
        
        r1 = p_dh.add_run()
        r1.text = f"{dtitle}  "
        r1.font.name = FONT_HEADING
        r1.font.size = Pt(10.5)
        r1.font.bold = True
        r1.font.color.rgb = dcol

        r2 = p_dh.add_run()
        r2.text = f"[{dsub}]"
        r2.font.name = FONT_MONO
        r2.font.size = Pt(8.5)
        r2.font.color.rgb = COLOR_TEXT_MUTED

        p_dd = tf.add_paragraph()
        p_dd.space_before = Pt(2)
        p_dd.text = ddesc
        p_dd.font.name = FONT_BODY
        p_dd.font.size = Pt(8.5)
        p_dd.font.color.rgb = COLOR_TEXT_SECONDARY

    # Right Column: Data Center Photo + Holistic Security Callout
    rx = Inches(8.4)
    rw = Inches(4.133)
    
    add_photo_card(
        slide,
        'assets/slide5_datacenter.jpg',
        rx, Inches(1.65), rw, Inches(3.3),
        caption="Figure 5.1: Enterprise server infrastructure enforcing multi-domain security policies.",
        border_color=COLOR_CYAN_ACCENT
    )

    hol_card = add_card(slide, rx, Inches(5.35), rw, Inches(1.5), bg_color=COLOR_CARD_HIGHLIGHT, border_color=COLOR_BLUE_ACCENT)
    tb_h = slide.shapes.add_textbox(rx + Inches(0.12), Inches(5.42), rw - Inches(0.24), Inches(1.35))
    tf_h = tb_h.text_frame
    tf_h.word_wrap = True
    p_hh = tf_h.paragraphs[0]
    p_hh.text = "THE WEAKEST LINK PRINCIPLE"
    p_hh.font.name = FONT_HEADING
    p_hh.font.size = Pt(10)
    p_hh.font.bold = True
    p_hh.font.color.rgb = COLOR_CYAN_LIGHT

    h_bullets = [
        "Systemic Interdependence: A military-grade Next-Gen Firewall (Network Security) is entirely useless if a developer commits an SQL injection vulnerability into a login script (Application Security).",
        "Defense-in-Depth: True enterprise resilience requires overlapping, non-redundant controls across all five domains simultaneously."
    ]
    for hb in h_bullets:
        p = tf_h.add_paragraph()
        p.space_before = Pt(2)
        p.text = f"• {hb}"
        p.font.name = FONT_BODY
        p.font.size = Pt(8.5)
        p.font.color.rgb = COLOR_TEXT_SECONDARY

    # Speaker Notes
    set_speaker_notes(slide, """
SLIDE 5: TYPES OF CYBERSECURITY (SPEAKER SCRIPT)

• CORE TALKING POINTS:
1. "When laymen think of cybersecurity, they usually imagine a firewall blocking hackers. But in professional computer science engineering, cybersecurity is organized into five distinct technical disciplines."
2. "Let's review these five disciplines:
   - Network Security: Protects the plumbing—traffic passing through switches, routers, and firewalls.
   - Application Security: Crucial for software engineers! This is where we write clean, secure code. We run SAST and DAST scans to catch buffer overflows, SQL injections, and broken authentication before code reaches production.
   - Endpoint Security: Focuses on the devices—employee laptops, mobile phones, and servers. Even if an employee downloads a malicious file, an EDR agent detects suspicious PowerShell or process-hollowing activity and isolates the host.
   - Cloud Security: In the cloud era, we operate under the Cloud Shared Responsibility Model. AWS or Microsoft secures the physical data centers, but the customer is responsible for configuring IAM roles, bucket permissions, and container security.
   - Data Security: The final line of defense. Even if an attacker penetrates the network and server, if the underlying database is strongly encrypted with AES-256 and keys are locked in an HSM, the stolen data is useless gibberish."
3. "Emphasize the 'Weakest Link Principle' on the bottom right: Security is like a chain. The adversary never attacks where you are strongest; they systematically scan for the single weakest link."

• TRANSITION:
"Now that we have established the broad landscape of cybersecurity, we enter the mathematical engine that powers all modern digital defense: Cryptography."
""")

print("Part 1 (Slides 1 to 5) compiled successfully.")
