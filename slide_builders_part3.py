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
# SLIDE 11: AUTHENTICATION & ACCESS CONTROL
# ==============================================================================
def build_slide_11(prs):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    apply_background(slide)
    add_header(
        slide,
        badge_text="MODULE 03: IDENTITY & DEFENSE",
        title_text="Identity, Authentication & Zero Trust Architecture",
        subtitle_text="Enforcing granular access control based on continuous verification rather than perimeter trust",
        slide_num=11
    )

    left_w = Inches(7.3)
    
    # Card 1: AAA Framework & The 3 Factors of MFA
    c1 = add_card(slide, Inches(0.8), Inches(1.65), left_w, Inches(2.5), border_color=COLOR_CYAN_ACCENT)
    tb1 = slide.shapes.add_textbox(Inches(0.95), Inches(1.72), left_w - Inches(0.3), Inches(2.35))
    tf1 = tb1.text_frame
    tf1.word_wrap = True
    p1 = tf1.paragraphs[0]
    p1.text = "1. Identity Fundamentals & The Three Pillars of MFA"
    p1.font.name = FONT_HEADING
    p1.font.size = Pt(11)
    p1.font.bold = True
    p1.font.color.rgb = COLOR_CYAN_LIGHT

    aaa_items = [
        ("The AAA Framework:", "Identification (claiming identity via username), Authentication (proving identity cryptographically), and Authorization (evaluating granular permissions to specific resources)."),
        ("Factor 1: Something You Know:", "Traditional passwords, PINs, secret passphrases. Highly vulnerable to credential stuffing, phishing, keystroke loggers, and database dumps."),
        ("Factor 2: Something You Have:", "Hardware cryptographic tokens (FIDO2 / WebAuthn, YubiKeys), Time-based One-Time Passwords (TOTP via Google Authenticator). Note: SMS OTP is insecure due to SIM-swapping."),
        ("Factor 3: Something You Are:", "Biometric markers (fingerprint ridges, facial depth mapping, iris patterns) verified locally within tamper-proof hardware secure enclaves."),
        ("FIDO2 / WebAuthn Standard:", "Phishing-proof passwordless standard based on public-key cryptography. The browser cryptographically proves possession of a private key bound to the exact origin domain.")
    ]
    for lbl, desc in aaa_items:
        p = tf1.add_paragraph()
        p.space_before = Pt(2.5)
        r1 = p.add_run()
        r1.text = f"• {lbl} "
        r1.font.bold = True
        r1.font.size = Pt(8.6)
        r1.font.color.rgb = COLOR_BLUE_ACCENT
        r2 = p.add_run()
        r2.text = desc
        r2.font.size = Pt(8.6)
        r2.font.color.rgb = COLOR_TEXT_SECONDARY

    # Card 2: Zero Trust Architecture (NIST SP 800-207)
    c2 = add_card(slide, Inches(0.8), Inches(4.3), left_w, Inches(2.5), bg_color=COLOR_CARD_HIGHLIGHT, border_color=COLOR_GREEN_ACCENT)
    tb2 = slide.shapes.add_textbox(Inches(0.95), Inches(4.37), left_w - Inches(0.3), Inches(2.35))
    tf2 = tb2.text_frame
    tf2.word_wrap = True
    p2 = tf2.paragraphs[0]
    p2.text = "2. Zero Trust Architecture (NIST SP 800-207): Paradigm Shift"
    p2.font.name = FONT_HEADING
    p2.font.size = Pt(11)
    p2.font.bold = True
    p2.font.color.rgb = COLOR_GREEN_ACCENT

    zt_points = [
        ("Core Maxim: 'Never Trust, Always Verify':", "Assume the internal corporate network is already fully compromised by adversaries. Traditional castle-and-moat perimeter models are dead."),
        ("Explicit Context-Based Verification:", "Every single access request is independently authenticated and authorized using user identity, device posture health, geolocation, and real-time behavioral risk scores."),
        ("Micro-segmentation:", "Dividing enterprise networks into granular zones. Prevents adversaries from moving laterally if a single workstation or server is breached."),
        ("Principle of Least Privilege (PoLP):", "Granting users and service accounts only the absolute minimum permissions and ephemeral access tokens (Just-In-Time access) necessary to perform their duty.")
    ]
    for lbl, desc in zt_points:
        p = tf2.add_paragraph()
        p.space_before = Pt(3)
        r1 = p.add_run()
        r1.text = f"• {lbl} "
        r1.font.bold = True
        r1.font.size = Pt(8.6)
        r1.font.color.rgb = COLOR_CYAN_LIGHT
        r2 = p.add_run()
        r2.text = desc
        r2.font.size = Pt(8.6)
        r2.font.color.rgb = COLOR_TEXT_PRIMARY

    # ---------------------------------------------------------
    # RIGHT COLUMN: 2 PHOTOS (Biometrics & MFA Hardware)
    # ---------------------------------------------------------
    rx = Inches(8.4)
    rw = Inches(4.133)
    
    # Photo 1: Biometric Access
    add_photo_card(
        slide,
        'assets/slide11_auth.jpg',
        rx, Inches(1.65), rw, Inches(2.35),
        caption="Figure 11.1: Localized biometric scanner verifying biometric hashes in silicon.",
        border_color=COLOR_CYAN_ACCENT
    )

    # Photo 2: Hardware MFA
    add_photo_card(
        slide,
        'assets/slide11_mfa.jpg',
        rx, Inches(4.5), rw, Inches(2.3),
        caption="Figure 11.2: FIDO2 hardware cryptographic security key preventing phishing.",
        border_color=COLOR_GREEN_ACCENT
    )

    # Speaker Notes
    set_speaker_notes(slide, """
SLIDE 11: AUTHENTICATION & ACCESS CONTROL (SPEAKER SCRIPT)

• CORE TALKING POINTS:
1. "In security engineering, we make a clear distinction in the AAA model:
   - Identification is who you say you are.
   - Authentication is proving it.
   - Authorization is what you are permitted to do once inside."
2. "Why are passwords dead? Because humans choose bad passwords, reuse them across 10 websites, and easily fall for phishing emails. Multi-Factor Authentication requires at least two out of three distinct factors:
   - Something you know (password)
   - Something you have (hardware key or authenticator app)
   - Something you are (biometrics)"
3. "Pay attention to FIDO2 / WebAuthn: Traditional SMS OTP is vulnerable to SIM-swap attacks, and basic 2FA codes can still be intercepted by reverse-proxy phishing kits (like Evilginx). But FIDO2 uses public-key cryptography bound to the website's domain in the browser. Even if a user falls for a fake phishing site, the hardware key will NOT send the cryptographic assertion because the domain does not match!"
4. "Finally, Zero Trust Architecture: The old 'castle-and-moat' model assumed that once someone was inside the office Wi-Fi, they were trusted. In Zero Trust, we assume the breach has already occurred. Every packet and every API call must prove identity and health."

• TRANSITION:
"What happens when organizations fail to enforce these exact principles? Let us analyze three real-world catastrophic cyber breaches."
""")


# ==============================================================================
# SLIDE 12: REAL-WORLD CASE STUDIES
# ==============================================================================
def build_slide_12(prs):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    apply_background(slide)
    add_header(
        slide,
        badge_text="MODULE 03: POST-MORTEM ANALYSIS",
        title_text="Dissecting Real-World Cyber Catastrophes",
        subtitle_text="Engineering lessons extracted from historical software vulnerabilities and systemic security breakdowns",
        slide_num=12
    )

    left_w = Inches(7.3)
    c_h = Inches(1.58)
    gap_y = Inches(0.12)
    start_y = Inches(1.65)

    cases = [
        (
            "Case 01: WannaCry Ransomware Pandemic (May 2017)",
            COLOR_RED_ACCENT,
            "Exploit Vector: Weaponized EternalBlue (MS17-010), a stolen NSA exploit targeting a critical remote code execution buffer overflow in Windows legacy SMBv1 protocol.\n"
            "Impact: Infected 200,000+ endpoints across 150 nations within hours. Paralyzed the UK National Health Service (NHS), grounding ambulances, cancelling surgeries, and impacting global shipping giants (FedEx, Maersk).\n"
            "Root Cause: Failure to apply an official Microsoft patch issued two months prior; lack of network segmentation."
        ),
        (
            "Case 02: SolarWinds Supply Chain Compromise (2020)",
            COLOR_AMBER_ACCENT,
            "Exploit Vector: Highly sophisticated nation-state APT compromised the internal build pipeline of SolarWinds Orion software, injecting the SUNBURST Trojan backdoor into cryptographically signed updates.\n"
            "Impact: 18,000 corporate and US federal government networks unwittingly downloaded the trojanized update; adversaries surveilled critical systems undetected for over 9 months.\n"
            "Root Cause: Blind trust in software updates; lack of Software Bill of Materials (SBOM) and build integrity validation."
        ),
        (
            "Case 03: Colonial Pipeline Ransomware Shutdown (May 2021)",
            COLOR_CYAN_LIGHT,
            "Exploit Vector: DarkSide ransomware gang breached the corporate virtual private network (VPN) using a single compromised employee password discovered in an unindexed dark web dump.\n"
            "Impact: Colonial Pipeline proactively shut down 5,500 miles of fuel pipelines delivering 45% of East Coast US gasoline, causing panic buying and triggering a regional state of national emergency.\n"
            "Root Cause: Lack of Multi-Factor Authentication (MFA) on legacy VPN endpoints; un-decommissioned inactive accounts."
        )
    ]

    for idx, (ctitle, ccol, cdesc) in enumerate(cases):
        cy = start_y + idx * (c_h + gap_y)
        c = add_card(slide, Inches(0.8), cy, left_w, c_h, border_color=COLOR_CARD_BORDER)
        
        strip = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0.8), cy, Inches(0.08), c_h)
        strip.fill.solid()
        strip.fill.fore_color.rgb = ccol
        strip.line.fill.background()

        tb = slide.shapes.add_textbox(Inches(1.0), cy + Inches(0.05), left_w - Inches(0.3), c_h - Inches(0.1))
        tf = tb.text_frame
        tf.word_wrap = True
        
        p = tf.paragraphs[0]
        p.text = ctitle
        p.font.name = FONT_HEADING
        p.font.size = Pt(10)
        p.font.bold = True
        p.font.color.rgb = ccol

        for line in cdesc.split('\n'):
            p_line = tf.add_paragraph()
            p_line.space_before = Pt(1.5)
            parts = line.split(': ')
            r1 = p_line.add_run()
            r1.text = f"• {parts[0]}: "
            r1.font.bold = True
            r1.font.size = Pt(8.2)
            r1.font.color.rgb = COLOR_CYAN_LIGHT
            if len(parts) > 1:
                r2 = p_line.add_run()
                r2.text = parts[1]
                r2.font.size = Pt(8.2)
                r2.font.color.rgb = COLOR_TEXT_SECONDARY

    # ---------------------------------------------------------
    # RIGHT COLUMN: FORENSICS PHOTO & ENGINEERING TAKEAWAY
    # ---------------------------------------------------------
    rx = Inches(8.4)
    rw = Inches(4.133)
    
    add_photo_card(
        slide,
        'assets/slide12_breach.jpg',
        rx, start_y, rw, Inches(3.2),
        caption="Figure 12.1: Digital forensics & incident response unit analyzing cyber breach telemetry.",
        border_color=COLOR_RED_ACCENT
    )

    les_card = add_card(slide, rx, Inches(5.2), rw, Inches(1.6), bg_color=COLOR_INNER_BOX, border_color=COLOR_AMBER_ACCENT)
    tb_l = slide.shapes.add_textbox(rx + Inches(0.12), Inches(5.26), rw - Inches(0.24), Inches(1.48))
    tf_l = tb_l.text_frame
    tf_l.word_wrap = True
    p_lh = tf_l.paragraphs[0]
    p_lh.text = "CRUCIAL LESSON FOR CS ENGINEERS"
    p_lh.font.name = FONT_HEADING
    p_lh.font.size = Pt(9.5)
    p_lh.font.bold = True
    p_lh.font.color.rgb = COLOR_AMBER_ACCENT

    lessons = [
        "Cryptography Was Never Broken: In none of these catastrophic attacks was AES or RSA cracked. The math remained completely secure.",
        "Systemic Failure: Breaches occurred because of unpatched software, single-factor logins, and compromised build pipelines.",
        "Takeaway: Cryptography without rigorous engineering hygiene and zero-trust policies is merely an illusion of security."
    ]
    for ls in lessons:
        p = tf_l.add_paragraph()
        p.space_before = Pt(2)
        p.text = f"✔ {ls}"
        p.font.name = FONT_BODY
        p.font.size = Pt(8.2)
        p.font.color.rgb = COLOR_TEXT_PRIMARY

    # Speaker Notes
    set_speaker_notes(slide, """
SLIDE 12: REAL-WORLD CASE STUDIES (SPEAKER SCRIPT)

• CORE TALKING POINTS:
1. "As computer science engineers, analyzing historical failures is how we learn to design robust systems. Let's examine three milestone attacks."
2. "Case 1: WannaCry (2017). A buffer overflow in SMBv1. Microsoft had released the security patch MS17-010 two months earlier, but hospitals and corporations failed to patch their servers. WannaCry scanned port 445, self-propagated automatically, and locked computers with AES and RSA. It was stopped only when a researcher discovered a hardcoded kill-switch domain."
3. "Case 2: SolarWinds (2020). This represents the nightmare of Supply Chain Attacks. Hackers didn't attack government databases directly; they compromised the build system of an IT monitoring vendor. The malware was signed with a legitimate digital certificate and pushed as an official update! That is why we now mandate Software Bill of Materials (SBOM) and hermetic builds."
4. "Case 3: Colonial Pipeline (2021). The largest fuel pipeline in the United States was shut down not by an advanced zero-day, but by a SINGLE password with NO Multi-Factor Authentication. A classic example of an unmanaged legacy account."
5. "Emphasize the bottom right box: In all three cases, the cryptographic math held firm. AES was not cracked. The human and operational systems failed."

• TRANSITION:
"Now, what does the future hold? Two revolutionary computational forces are colliding with cybersecurity: Quantum Computing and Artificial Intelligence."
""")


# ==============================================================================
# SLIDE 13: QUANTUM THREATS & AI IN CYBERSECURITY
# ==============================================================================
def build_slide_13(prs):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    apply_background(slide)
    add_header(
        slide,
        badge_text="MODULE 03: EMERGING FRONTIERS",
        title_text="The Next Frontier: Quantum Supremacy & Artificial Intelligence",
        subtitle_text="Preparing computational architectures for quantum cryptanalysis and autonomous adversarial AI",
        slide_num=13
    )

    half_w = Inches(5.7)
    card_h = Inches(5.15)
    start_y = Inches(1.65)

    # ---------------------------------------------------------
    # LEFT COLUMN: THE QUANTUM CRYPTOGRAPHIC THREAT & PQC
    # ---------------------------------------------------------
    c_q = add_card(slide, Inches(0.8), start_y, half_w, card_h, border_color=COLOR_CYAN_ACCENT)
    tb_q = slide.shapes.add_textbox(Inches(0.95), start_y + Inches(0.12), half_w - Inches(0.3), card_h - Inches(0.24))
    tf_q = tb_q.text_frame
    tf_q.word_wrap = True

    p_qh = tf_q.paragraphs[0]
    p_qh.text = "1. THE QUANTUM THREAT & POST-QUANTUM CRYPTOGRAPHY"
    p_qh.font.name = FONT_HEADING
    p_qh.font.size = Pt(10.5)
    p_qh.font.bold = True
    p_qh.font.color.rgb = COLOR_CYAN_LIGHT

    q_points = [
        ("The Quantum Paradigm:", "Quantum computers leverage quantum bits (qubits), superposition, and entanglement to solve specific mathematical classes exponentially faster than classical Turing machines."),
        ("Shor's Algorithm (1994):", "Solves prime factorization and discrete logarithms in polynomial time O((log N)^3). Completely destroys RSA, ECC, and Diffie-Hellman once cryptographically relevant quantum computers (CRQC) emerge."),
        ("Grover's Algorithm:", "Accelerates unstructured database search with quadratic speedup O(sqrt(N)). Halves symmetric key security: AES-128 drops to 64-bit (insecure); AES-256 remains resilient with 128-bit quantum strength."),
        ("'Harvest Now, Decrypt Later':", "Adversaries currently intercept and store encrypted petabytes of government and corporate data to decrypt them retroactively when quantum systems become available."),
        ("NIST Post-Quantum Standards (2024):", "ML-KEM (CRYSTALS-Kyber) for lattice-based key encapsulation; ML-DSA (CRYSTALS-Dilithium) for digital signatures.")
    ]
    for lbl, desc in q_points:
        p = tf_q.add_paragraph()
        p.space_before = Pt(2.5)
        r1 = p.add_run()
        r1.text = f"• {lbl} "
        r1.font.bold = True
        r1.font.size = Pt(8.5)
        r1.font.color.rgb = COLOR_CYAN_ACCENT
        r2 = p.add_run()
        r2.text = desc
        r2.font.size = Pt(8.5)
        r2.font.color.rgb = COLOR_TEXT_SECONDARY

    # Left Photo: Quantum Processor
    add_photo_card(
        slide,
        'assets/slide13_quantum.jpg',
        Inches(0.95), Inches(5.35), half_w - Inches(0.3), Inches(1.3),
        caption="Figure 13.1: Cryogenic quantum processor threatening classical public-key cryptography.",
        border_color=COLOR_CYAN_ACCENT
    )

    # ---------------------------------------------------------
    # RIGHT COLUMN: AI IN CYBERSECURITY (THE DUAL-USE ARMS RACE)
    # ---------------------------------------------------------
    rx = Inches(6.8)
    c_ai = add_card(slide, rx, start_y, half_w, card_h, border_color=COLOR_PURPLE_ACCENT)
    tb_ai = slide.shapes.add_textbox(rx + Inches(0.15), start_y + Inches(0.12), half_w - Inches(0.3), card_h - Inches(0.24))
    tf_ai = tb_ai.text_frame
    tf_ai.word_wrap = True

    p_aih = tf_ai.paragraphs[0]
    p_aih.text = "2. ARTIFICIAL INTELLIGENCE: DUAL-USE SECURITY ARMS RACE"
    p_aih.font.name = FONT_HEADING
    p_aih.font.size = Pt(10.5)
    p_aih.font.bold = True
    p_aih.font.color.rgb = COLOR_PURPLE_ACCENT

    ai_points = [
        ("Offensive AI (Adversarial Capabilities):", ""),
        ("  • Hyper-Targeted Phishing:", "LLMs autonomously analyzing executive social profiles to draft un-detectable, personalized spear-phishing lures at scale."),
        ("  • Polymorphic Malware Mutation:", "AI models dynamically rewriting malicious code to evade signature and static heuristic antivirus detection engines."),
        ("  • Automated Vulnerability Discovery:", "Autonomous agents fuzzing binaries 24/7 to discover and synthesize zero-day exploits before defenders can patch them."),
        ("Defensive AI (Cyber Defense Resiliency):", ""),
        ("  • Behavioral Anomaly Detection:", "Machine learning models ingesting billions of network telemetry logs per second to spot subtle lateral movement."),
        ("  • Automated SOC Triage (SOAR):", "AI agents parsing, contextualizing, and containing threat alerts in milliseconds, relieving human burnout.")
    ]
    for lbl, desc in ai_points:
        p = tf_ai.add_paragraph()
        p.space_before = Pt(2)
        r1 = p.add_run()
        r1.text = f"{lbl} "
        r1.font.bold = True
        r1.font.size = Pt(8.5)
        r1.font.color.rgb = COLOR_RED_ACCENT if "Offensive" in lbl else (COLOR_GREEN_ACCENT if "Defensive" in lbl else COLOR_CYAN_LIGHT)
        r2 = p.add_run()
        r2.text = desc
        r2.font.size = Pt(8.5)
        r2.font.color.rgb = COLOR_TEXT_SECONDARY

    # Right Photo: AI Neural Network
    add_photo_card(
        slide,
        'assets/slide13_ai.jpg',
        rx + Inches(0.15), Inches(5.35), half_w - Inches(0.3), Inches(1.3),
        caption="Figure 13.2: Deep neural networks powering autonomous threat detection and behavioral analysis.",
        border_color=COLOR_PURPLE_ACCENT
    )

    # Speaker Notes
    set_speaker_notes(slide, """
SLIDE 13: QUANTUM THREATS & AI IN CYBERSECURITY (SPEAKER SCRIPT)

• CORE TALKING POINTS:
1. "We stand at the threshold of two transformative technologies that will redefine cybersecurity: Quantum Computing and Artificial Intelligence."
2. "Let's demystify the Quantum Threat on the left:
   - Today, our entire internet security (RSA, ECC, Diffie-Hellman) depends on math problems that classical supercomputers take billions of years to solve—specifically factoring large prime numbers.
   - But in 1994, Peter Shor proved that a quantum computer can factor numbers in polynomial time using Shor's Algorithm. When large-scale quantum computers arrive, RSA and ECC will be broken instantly.
   - Notice the threat: 'Harvest Now, Decrypt Later'. Hostile intelligence agencies are recording encrypted TLS data today, waiting for quantum computers to decrypt it in the 2030s!
   - That is why NIST just released official Post-Quantum Cryptography standards: ML-KEM and ML-DSA, which use multi-dimensional lattice mathematics that quantum computers cannot crack."
3. "Now look at the right side: The AI Security Arms Race.
   - Adversaries are using generative AI to write flawless spear-phishing emails in 50 languages and mutate malware code on the fly.
   - On the defensive side, human analysts cannot analyze 10 billion telemetry logs per second. We need defensive machine learning to catch anomalous user behavior and isolate compromised laptops within milliseconds."

• TRANSITION:
"Faced with these formidable threats, how do we engineer a truly resilient defense? Let us explore Defense-in-Depth and Secure Engineering Best Practices."
""")


# ==============================================================================
# SLIDE 14: DEFENSE-IN-DEPTH & BEST PRACTICES
# ==============================================================================
def build_slide_14(prs):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    apply_background(slide)
    add_header(
        slide,
        badge_text="MODULE 03: DEFENSIVE ARCHITECTURE",
        title_text="Defense-in-Depth & Secure Software Engineering",
        subtitle_text="Architecting multi-layered defensive moats and embedding security into every phase of development",
        slide_num=14
    )

    left_w = Inches(7.3)
    
    # Card 1: The Concentric 5-Layer Moat
    c1 = add_card(slide, Inches(0.8), Inches(1.65), left_w, Inches(2.55), border_color=COLOR_CYAN_ACCENT)
    tb1 = slide.shapes.add_textbox(Inches(0.95), Inches(1.72), left_w - Inches(0.3), Inches(2.4))
    tf1 = tb1.text_frame
    tf1.word_wrap = True
    p1 = tf1.paragraphs[0]
    p1.text = "1. Concentric Defense-in-Depth Moat (Layered Controls)"
    p1.font.name = FONT_HEADING
    p1.font.size = Pt(11)
    p1.font.bold = True
    p1.font.color.rgb = COLOR_CYAN_LIGHT

    layers = [
        ("Layer 1: Perimeter / Edge:", "Cloud DDoS shielding (Anycast), Web Application Firewalls (WAF), secure DNS filtering, and VPN gateways."),
        ("Layer 2: Network Infrastructure:", "VLAN micro-segmentation, Zero Trust Network Access (ZTNA), Network Intrusion Prevention (NIPS), and encrypted tunnels."),
        ("Layer 3: Host / Compute:", "OS kernel hardening, Endpoint Detection and Response (EDR), automated patch orchestration, and disabled unused ports."),
        ("Layer 4: Application Logic:", "Secure code reviews, input validation / prepared SQL statements, container image vulnerability scanning, and rate limiting."),
        ("Layer 5: Data Core:", "AES-256 cryptographic encryption at-rest, TLS 1.3 in-transit, Data Loss Prevention (DLP), and immutable backups.")
    ]
    for lbl, desc in layers:
        p = tf1.add_paragraph()
        p.space_before = Pt(2.5)
        r1 = p.add_run()
        r1.text = f"• {lbl} "
        r1.font.bold = True
        r1.font.size = Pt(8.6)
        r1.font.color.rgb = COLOR_BLUE_ACCENT
        r2 = p.add_run()
        r2.text = desc
        r2.font.size = Pt(8.6)
        r2.font.color.rgb = COLOR_TEXT_SECONDARY

    # Card 2: 4 Engineering Best Practices
    c2 = add_card(slide, Inches(0.8), Inches(4.35), left_w, Inches(2.45), bg_color=COLOR_CARD_HIGHLIGHT, border_color=COLOR_GREEN_ACCENT)
    tb2 = slide.shapes.add_textbox(Inches(0.95), Inches(4.42), left_w - Inches(0.3), Inches(2.3))
    tf2 = tb2.text_frame
    tf2.word_wrap = True
    p2 = tf2.paragraphs[0]
    p2.text = "2. Core Engineering Best Practices for Developers"
    p2.font.name = FONT_HEADING
    p2.font.size = Pt(11)
    p2.font.bold = True
    p2.font.color.rgb = COLOR_GREEN_ACCENT

    practices = [
        ("DevSecOps & 'Shift-Left' Security:", "Integrate security automation directly into CI/CD pipelines (SAST, DAST, dependency checks). Catch flaws in code, not in production. NEVER hardcode API secrets in Git!"),
        ("Principle of Least Privilege (PoLP):", "Grant every user, service, and cloud container only the minimum access permissions strictly required to perform their discrete function."),
        ("Immutable 3-2-1 Backup Strategy:", "Maintain 3 copies of vital data, on 2 different physical media, with 1 copy stored completely off-site and air-gapped/write-once (WORM) to defeat ransomware."),
        ("Incident Response Lifecycle (NIST SP 800-61):", "Formal 4-phase cycle: Preparation -> Detection & Analysis -> Containment, Eradication, & Recovery -> Post-Incident Lessons Learned.")
    ]
    for lbl, desc in practices:
        p = tf2.add_paragraph()
        p.space_before = Pt(2.5)
        r1 = p.add_run()
        r1.text = f"• {lbl} "
        r1.font.bold = True
        r1.font.size = Pt(8.6)
        r1.font.color.rgb = COLOR_CYAN_LIGHT
        r2 = p.add_run()
        r2.text = desc
        r2.font.size = Pt(8.6)
        r2.font.color.rgb = COLOR_TEXT_PRIMARY

    # ---------------------------------------------------------
    # RIGHT COLUMN: SOC OPERATIONS PHOTO & SECURITY METRIC
    # ---------------------------------------------------------
    rx = Inches(8.4)
    rw = Inches(4.133)
    
    add_photo_card(
        slide,
        'assets/slide14_defense.jpg',
        rx, Inches(1.65), rw, Inches(3.3),
        caption="Figure 14.1: Cybersecurity command center executing defense-in-depth protocols.",
        border_color=COLOR_CYAN_ACCENT
    )

    tip_card = add_card(slide, rx, Inches(5.35), rw, Inches(1.45), bg_color=COLOR_INNER_BOX, border_color=COLOR_BLUE_ACCENT)
    tb_t = slide.shapes.add_textbox(rx + Inches(0.12), Inches(5.42), rw - Inches(0.24), Inches(1.3))
    tf_t = tb_t.text_frame
    tf_t.word_wrap = True
    p_th = tf_t.paragraphs[0]
    p_th.text = "THE 'SHIFT-LEFT' ECONOMIC REALITY"
    p_th.font.name = FONT_HEADING
    p_th.font.size = Pt(9.5)
    p_th.font.bold = True
    p_th.font.color.rgb = COLOR_CYAN_LIGHT

    shift_text = [
        "Architectural Fix: Remediating a vulnerability during the design and coding phase costs $80 - $200.",
        "Post-Deployment Breach: Mitigating the exact same flaw after an exploit in production costs over $4.45 Million in average incident damages, fines, and forensic audits."
    ]
    for st in shift_text:
        p = tf_t.add_paragraph()
        p.space_before = Pt(2)
        p.text = f"✔ {st}"
        p.font.name = FONT_BODY
        p.font.size = Pt(8.2)
        p.font.color.rgb = COLOR_TEXT_PRIMARY

    # Speaker Notes
    set_speaker_notes(slide, """
SLIDE 14: DEFENSE-IN-DEPTH & BEST PRACTICES (SPEAKER SCRIPT)

• CORE TALKING POINTS:
1. "The medieval castle is the classic metaphor for Defense-in-Depth. A castle didn't just have one wall; it had a moat, a drawbridge, outer ramparts, inner walls, and a fortified keep. If an enemy breached the moat, the outer wall stopped them."
2. "In modern computing, we layer security from the Perimeter down to the Data Core:
   - Perimeter: Firewalls and DDoS shields.
   - Network: Segmenting development servers away from finance servers.
   - Host: EDR agents detecting suspicious processes.
   - Application: Writing clean, secure code that sanitizes input.
   - Data: Encrypting the database so that even if all outer layers fall, the raw files are unreadable."
3. "For student software engineers, pay close attention to 'Shift-Left' Security:
   - Traditionally, developers wrote code, shipped it, and let security teams worry about it later.
   - In modern DevSecOps, security tests run automatically inside the GitHub/GitLab CI pipeline on every single commit.
   - Look at the economic metric on the right: Catching a bug during coding costs $100. Fixing it after a public breach costs millions!"
4. "Remember the 3-2-1 backup rule: It is the single most effective antidote against ransomware."

• TRANSITION:
"Now, let us synthesize everything we have covered, summarize our key engineering takeaways, and open the floor for discussion."
""")


# ==============================================================================
# SLIDE 15: CONCLUSION, SUMMARY & Q&A
# ==============================================================================
def build_slide_15(prs):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    apply_background(slide)
    add_header(
        slide,
        badge_text="SEMINAR CONCLUSION & SYNTHESIS",
        title_text="Conclusion, Future Outlook & Interactive Q&A",
        subtitle_text="Synthesizing cryptographic rigor with systemic cybersecurity to protect the global digital society",
        slide_num=15
    )

    left_w = Inches(7.3)
    
    # Card 1: 4 Core Takeaways for CS Engineers
    c1 = add_card(slide, Inches(0.8), Inches(1.65), left_w, Inches(3.2), border_color=COLOR_CYAN_ACCENT)
    tb1 = slide.shapes.add_textbox(Inches(0.95), Inches(1.72), left_w - Inches(0.3), Inches(3.05))
    tf1 = tb1.text_frame
    tf1.word_wrap = True
    p1 = tf1.paragraphs[0]
    p1.text = "Core Seminar Synthesis: Key Engineering Takeaways"
    p1.font.name = FONT_HEADING
    p1.font.size = Pt(11.5)
    p1.font.bold = True
    p1.font.color.rgb = COLOR_CYAN_LIGHT

    takeaways = [
        ("1. Security is a Lifecycle, Not an Add-On:", "Cryptography provides the mathematical certainty, but systemic policies, human discipline, and architectural hygiene dictate true organizational defense. Both disciplines are inextricably intertwined."),
        ("2. Cryptography is Foundational Infrastructure:", "From TLS 1.3 web tunnels and SSH administration to blockchain ledgers and Signal E2EE—every modern software system depends wholly on symmetric, asymmetric, and hash primitives."),
        ("3. Embrace Zero Trust & Defense-in-Depth:", "Never assume the internal network is safe. Enforce hardware MFA (FIDO2), least privilege access, micro-segmentation, and immutable offline backups as non-negotiable baselines."),
        ("4. Prepare for the Post-Quantum & AI Horizon:", "Engineering students must design cryptographically agile systems capable of dropping in NIST Post-Quantum standards (ML-KEM/ML-DSA) and leverage defensive AI to counter automated threats.")
    ]
    for lbl, desc in takeaways:
        p = tf1.add_paragraph()
        p.space_before = Pt(3.5)
        r1 = p.add_run()
        r1.text = f"• {lbl} "
        r1.font.bold = True
        r1.font.size = Pt(8.8)
        r1.font.color.rgb = COLOR_BLUE_ACCENT
        r2 = p.add_run()
        r2.text = desc
        r2.font.size = Pt(8.8)
        r2.font.color.rgb = COLOR_TEXT_SECONDARY

    # Card 2: Standards & References Box
    c2 = add_card(slide, Inches(0.8), Inches(4.95), left_w, Inches(1.85), bg_color=COLOR_INNER_BOX, border_color=COLOR_CARD_BORDER)
    tb2 = slide.shapes.add_textbox(Inches(0.95), Inches(5.02), left_w - Inches(0.3), Inches(1.7))
    tf2 = tb2.text_frame
    tf2.word_wrap = True
    p2 = tf2.paragraphs[0]
    p2.text = "AUTHORITATIVE STANDARDS & ACADEMIC REFERENCES"
    p2.font.name = FONT_HEADING
    p2.font.size = Pt(10)
    p2.font.bold = True
    p2.font.color.rgb = COLOR_GREEN_ACCENT

    refs = [
        ("NIST Special Publications:", "SP 800-207 (Zero Trust Architecture), SP 800-61 (Incident Handling), FIPS 197 (AES)."),
        ("IETF RFC Standards:", "RFC 8446 (The Transport Layer Security Protocol TLS 1.3), RFC 4253 (SSH Protocol)."),
        ("Foundational Literature:", "William Stallings, 'Cryptography and Network Security: Principles and Practice', 8th Edition."),
        ("Industry Frameworks:", "OWASP Top 10 Web Application Vulnerabilities (2025/2026); MITRE ATT&CK Framework.")
    ]
    for rlbl, rdesc in refs:
        p = tf2.add_paragraph()
        p.space_before = Pt(2)
        r1 = p.add_run()
        r1.text = f"• {rlbl} "
        r1.font.bold = True
        r1.font.size = Pt(8.2)
        r1.font.color.rgb = COLOR_CYAN_LIGHT
        r2 = p.add_run()
        r2.text = rdesc
        r2.font.size = Pt(8.2)
        r2.font.color.rgb = COLOR_TEXT_PRIMARY

    # ---------------------------------------------------------
    # RIGHT COLUMN: SEMINAR AUDITORIUM PHOTO & Q&A CALLOUT
    # ---------------------------------------------------------
    rx = Inches(8.4)
    rw = Inches(4.133)
    
    add_photo_card(
        slide,
        'assets/slide15_seminar.jpg',
        rx, Inches(1.65), rw, Inches(3.2),
        caption="Figure 15.1: Academic technical seminar session and open technical dialogue.",
        border_color=COLOR_CYAN_ACCENT
    )

    qa_card = add_card(slide, rx, Inches(5.2), rw, Inches(1.6), bg_color=COLOR_CARD_HIGHLIGHT, border_color=COLOR_CYAN_ACCENT)
    tb_qa = slide.shapes.add_textbox(rx + Inches(0.12), Inches(5.28), rw - Inches(0.24), Inches(1.45))
    tf_qa = tb_qa.text_frame
    tf_qa.word_wrap = True
    p_qah = tf_qa.paragraphs[0]
    p_qah.text = "QUESTIONS & SEMINAR DISCUSSION"
    p_qah.alignment = PP_ALIGN.CENTER
    p_qah.font.name = FONT_HEADING
    p_qah.font.size = Pt(12)
    p_qah.font.bold = True
    p_qah.font.color.rgb = COLOR_CYAN_LIGHT

    p_qat = tf_qa.add_paragraph()
    p_qat.space_before = Pt(4)
    p_qat.text = "Thank you for your time and engagement!\nThe floor is now open for questions, technical discussion, and faculty feedback."
    p_qat.alignment = PP_ALIGN.CENTER
    p_qat.font.name = FONT_BODY
    p_qat.font.size = Pt(9.5)
    p_qat.font.color.rgb = COLOR_TEXT_PRIMARY

    # Speaker Notes
    set_speaker_notes(slide, """
SLIDE 15: CONCLUSION, SUMMARY & Q&A (SPEAKER SCRIPT)

• CONCLUDING ADDRESS:
"Respected faculty evaluators and peers, to summarize our technical seminar today:
1. Cybersecurity and Cryptography are two sides of the same coin. Cryptography provides the unyielding mathematical foundation—the ciphers, hash algorithms, and digital signatures that guarantee Confidentiality, Integrity, and Non-Repudiation.
2. But Cybersecurity provides the systemic architecture—the identity systems, endpoint monitoring, defense-in-depth layers, and human operational hygiene that prevent mathematical algorithms from being bypassed in production.
3. As aspiring computer science engineers, we must write code with a security-first mindset: adopting DevSecOps, following Zero Trust principles, and preparing for the emerging post-quantum era."

• CLOSING & Q&A INVITATION:
"I would like to express my sincere gratitude to my faculty guide, the Department of Computer Science & Engineering, and everyone present today for your valuable time and attention.
I warmly welcome any questions, comments, or technical discussions."

• SAMPLE Q&A DEFENSE PREPARATION FOR THE PRESENTER:
1. Q: "Why can't we just use AES-512 instead of AES-256 for post-quantum defense?"
   A: "Grover's algorithm reduces the effective security of symmetric ciphers by half (square-root speedup). AES-256 thus provides 128 bits of security against a quantum adversary, which remains practically uncrackable. AES-512 is unnecessary and adds computational overhead."
2. Q: "What is the biggest vulnerability in modern cybersecurity?"
   A: "The human factor—phishing and social engineering account for over 80% of corporate breaches. That is why hardware-backed MFA (FIDO2) and Zero Trust architecture are critical."
""")

print("Part 3 (Slides 11 to 15) compiled successfully.")
