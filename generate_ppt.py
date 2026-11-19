import os
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.enum.shapes import MSO_SHAPE

def create_presentation():
    prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)
    blank_layout = prs.slide_layouts[6] # blank slide

    # Theme Colors
    BG_COLOR = RGBColor(15, 23, 42)        # Slate 900
    CARD_BG = RGBColor(30, 41, 59)         # Slate 800
    CARD_BORDER = RGBColor(51, 65, 85)     # Slate 700
    ACCENT_CYAN = RGBColor(56, 189, 248)   # Sky 400
    ACCENT_BLUE = RGBColor(37, 99, 235)    # Blue 600
    ACCENT_GREEN = RGBColor(52, 211, 153)  # Emerald 400
    TEXT_LIGHT = RGBColor(248, 250, 252)   # Slate 50
    TEXT_MUTED = RGBColor(148, 163, 184)   # Slate 400
    BADGE_BG = RGBColor(14, 116, 144)      # Cyan 700

    def set_slide_background(slide):
        bg = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, prs.slide_width, prs.slide_height)
        bg.fill.solid()
        bg.fill.fore_color.rgb = BG_COLOR
        bg.line.fill.background()
        return bg

    def add_header(slide, title_text, category_text="CYBERSECURITY & CRYPTOGRAPHY"):
        # Category pill
        cat_box = slide.shapes.add_textbox(Inches(0.8), Inches(0.4), Inches(8), Inches(0.4))
        tf = cat_box.text_frame
        tf.word_wrap = True
        p = tf.paragraphs[0]
        p.text = category_text.upper()
        p.font.size = Pt(10)
        p.font.bold = True
        p.font.color.rgb = ACCENT_CYAN

        # Title
        title_box = slide.shapes.add_textbox(Inches(0.8), Inches(0.75), Inches(11.7), Inches(0.8))
        tf2 = title_box.text_frame
        tf2.word_wrap = True
        p2 = tf2.paragraphs[0]
        p2.text = title_text
        p2.font.size = Pt(24)
        p2.font.bold = True
        p2.font.color.rgb = TEXT_LIGHT

    # --- SLIDE 1: Title Slide ---
    slide1 = prs.slides.add_slide(blank_layout)
    set_slide_background(slide1)
    
    cover_img_path = os.path.join(os.path.dirname(__file__), "cover_image.jpg")
    if os.path.exists(cover_img_path):
        slide1.shapes.add_picture(cover_img_path, 0, 0, width=prs.slide_width, height=prs.slide_height)

    # Sleek glassmorphism-style card on the left for title typography
    card_shape = slide1.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.8), Inches(1.1), Inches(6.4), Inches(5.3))
    card_shape.fill.solid()
    card_shape.fill.fore_color.rgb = RGBColor(10, 17, 32)
    card_shape.line.color.rgb = ACCENT_CYAN
    card_shape.line.width = Pt(1.5)
    
    tf = card_shape.text_frame
    tf.word_wrap = True
    tf.margin_left = Inches(0.4)
    tf.margin_right = Inches(0.4)
    tf.margin_top = Inches(0.4)
    tf.margin_bottom = Inches(0.4)
    
    p0 = tf.paragraphs[0]
    p0.text = "COMPREHENSIVE MASTERCLASS"
    p0.font.size = Pt(12)
    p0.font.bold = True
    p0.font.color.rgb = ACCENT_CYAN
    p0.space_after = Pt(10)

    p1 = tf.add_paragraph()
    p1.text = "Cybersecurity &\nCryptography Methods"
    p1.font.size = Pt(32)
    p1.font.bold = True
    p1.font.color.rgb = TEXT_LIGHT
    p1.space_after = Pt(10)

    p2 = tf.add_paragraph()
    p2.text = "Step-by-Step Technical Guide: Architectures, Algorithms & Implementation Workflows"
    p2.font.size = Pt(14)
    p2.font.color.rgb = TEXT_MUTED
    p2.space_after = Pt(16)

    p3 = tf.add_paragraph()
    p3.text = "Key Modules Covered:\n• CIA Triad & Threat Modeling (STRIDE)\n• Symmetric (AES) & Asymmetric (RSA)\n• Key Exchange (Diffie-Hellman) & Hashing (SHA-256)\n• Digital Signatures (PKI) & TLS 1.3 Handshake\n• Defense-in-Depth & Post-Quantum Cryptography"
    p3.font.size = Pt(11)
    p3.font.color.rgb = ACCENT_GREEN

    # Speaker notes
    slide1.notes_slide.notes_text_frame.text = "Welcome to this comprehensive technical presentation on Cybersecurity and Cryptography Methods. We will examine the core defensive pillars and dive step-by-step into how fundamental cryptographic algorithms operate."

    # Helper function for 2-column or step-card slides
    def create_step_slide(slide_title, category, overview_items, steps, notes_text):
        slide = prs.slides.add_slide(blank_layout)
        set_slide_background(slide)
        add_header(slide, slide_title, category)

        # Left Card (Overview / Key Concepts)
        left_card = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.8), Inches(1.7), Inches(3.8), Inches(5.2))
        left_card.fill.solid()
        left_card.fill.fore_color.rgb = CARD_BG
        left_card.line.color.rgb = CARD_BORDER

        tf_l = left_card.text_frame
        tf_l.word_wrap = True
        tf_l.margin_left = Inches(0.25)
        tf_l.margin_right = Inches(0.25)
        tf_l.margin_top = Inches(0.25)
        
        p = tf_l.paragraphs[0]
        p.text = "Core Principles & Concepts"
        p.font.size = Pt(14)
        p.font.bold = True
        p.font.color.rgb = ACCENT_CYAN
        p.space_after = Pt(10)

        for item in overview_items:
            pi = tf_l.add_paragraph()
            pi.text = f"• {item[0]}: "
            pi.font.bold = True
            pi.font.size = Pt(11)
            pi.font.color.rgb = TEXT_LIGHT
            pi.space_after = Pt(6)
            
            # Sub-desc
            run = pi.add_run()
            run.text = item[1]
            run.font.bold = False
            run.font.color.rgb = TEXT_MUTED

        # Right Steps Area (Container)
        num_steps = len(steps)
        step_w = Inches(7.7)
        # Calculate dynamic height per step card
        step_h = (Inches(5.2) - (num_steps - 1) * Inches(0.12)) / num_steps
        
        for idx, step in enumerate(steps):
            y_pos = Inches(1.7) + idx * (step_h + Inches(0.12))
            scard = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(4.8), y_pos, step_w, step_h)
            scard.fill.solid()
            scard.fill.fore_color.rgb = CARD_BG
            scard.line.color.rgb = CARD_BORDER

            tf_s = scard.text_frame
            tf_s.word_wrap = True
            tf_s.margin_left = Inches(0.2)
            tf_s.margin_right = Inches(0.2)
            tf_s.margin_top = Inches(0.12)
            tf_s.margin_bottom = Inches(0.12)

            ps = tf_s.paragraphs[0]
            ps.text = f"STEP {idx+1}: {step[0]}"
            ps.font.size = Pt(12)
            ps.font.bold = True
            ps.font.color.rgb = ACCENT_GREEN
            ps.space_after = Pt(2)

            ps_desc = tf_s.add_paragraph()
            ps_desc.text = step[1]
            ps_desc.font.size = Pt(10)
            ps_desc.font.color.rgb = TEXT_LIGHT

        slide.notes_slide.notes_text_frame.text = notes_text

    # --- SLIDE 2: Cybersecurity Foundations & Risk Assessment ---
    create_step_slide(
        "Cybersecurity Foundations: The CIA Triad & Risk Assessment",
        "Module 1: Defense Fundamentals",
        [
            ("Confidentiality", "Ensuring data is accessible strictly to authorized individuals using encryption and access controls."),
            ("Integrity", "Guarantees information is accurate, genuine, and protected against unauthorized tampering or modification."),
            ("Availability", "Ensures high uptime, resilience, and unhindered access to authorized users via redundancy and disaster recovery."),
            ("Accountability", "Authentication, authorization, and non-repudiation audit trails for all actions.")
        ],
        [
            ("Asset Identification & Categorization", "Catalogue all critical digital assets, confidential customer databases, proprietary source code, and key API endpoints."),
            ("Vulnerability Scanning & Penetration Testing", "Run automated vulnerability scanners (e.g. Nessus, Qualys) and manual penetration testing to identify misconfigurations."),
            ("Threat Modeling (STRIDE / DREAD)", "Identify potential vectors: Spoofing, Tampering, Repudiation, Information Disclosure, Denial of Service, Elevation of Privilege."),
            ("Risk Prioritization & Impact Scoring", "Calculate risk matrix: Risk = Likelihood × Impact. Rank vulnerabilities by CVSS score to guide resource allocation."),
            ("Security Controls & Continuous Auditing", "Deploy defense-in-depth controls (MFA, WAF, SIEM logging, EDR) and establish 24/7 continuous monitoring.")
        ],
        "Notes: Emphasize that Cybersecurity is not a one-time product but an ongoing lifecycle. Every defense begins with knowing your assets and modeling threats systematically."
    )

    # --- SLIDE 3: The Cyber Attack Lifecycle (Cyber Kill Chain) ---
    create_step_slide(
        "The Cyber Attack Lifecycle: Cyber Kill Chain vs. Defense",
        "Module 2: Threat Modeling",
        [
            ("Adversary Objective", "Threat actors execute multi-stage campaigns to penetrate networks, maintain persistence, and exfiltrate data."),
            ("Defensive Philosophy", "Break any single link in the kill chain to neutralize the attack before objectives are achieved."),
            ("MITRE ATT&CK", "Maps real-world adversary tactics, techniques, and procedures (TTPs) across every phase."),
            ("Proactive Stance", "Shift from reactive fire-fighting to proactive threat hunting and automated containment.")
        ],
        [
            ("Reconnaissance (Passive & Active)", "Attacker scans domain records, OSINT, and open ports. Defense: Attack surface management & dark web monitoring."),
            ("Weaponization & Delivery", "Exploit crafted into payload (e.g., infected PDF). Defense: Email filtering, DMARC/SPF, and secure web gateways."),
            ("Exploitation & Code Execution", "Vulnerability triggered on host system. Defense: Rapid patch management, memory protection (DEP/ASLR), and EDR."),
            ("Installation & Persistence", "Malware installs backdoors or registry run-keys. Defense: Endpoint detection & response, application whitelisting."),
            ("Command & Control (C2 Beaconing)", "Infected host dials out to adversary C2 server. Defense: DNS sinkholing, TLS inspection, and network egress blocks."),
            ("Actions on Objectives (Exfiltration/Ransom)", "Data encrypted or stolen. Defense: Immutable backups, Data Loss Prevention (DLP), and micro-segmentation.")
        ],
        "Notes: Walk through the 6 core stages. Emphasize that catching the attacker early during Delivery or Exploitation is significantly cheaper than responding after Actions on Objectives."
    )

    # --- SLIDE 4: Cryptography Core Architecture & Terminology ---
    create_step_slide(
        "Cryptography Landscape: Core Concepts & Kerckhoffs's Principle",
        "Module 3: Cryptography Foundations",
        [
            ("Plaintext & Ciphertext", "Plaintext (M) is human-readable data; Ciphertext (C) is the unreadable scrambled output."),
            ("Encryption / Decryption", "Encryption: C = E_k(M); Decryption: M = D_k(C). Controlled by mathematical cryptographic keys."),
            ("Kerckhoffs's Principle", "'A cryptosystem should be secure even if everything about the system, except the key, is public knowledge.'"),
            ("Cryptanalysis", "The study of breaking ciphers (Brute force, Chosen-Plaintext, Side-channel attacks).")
        ],
        [
            ("Algorithm Selection", "Select an open, peer-reviewed standard algorithm (e.g., AES-256, RSA-3072, ECC secp256r1). Never invent proprietary crypto."),
            ("Secure Key Generation (CSPRNG)", "Generate cryptographic keys using a Cryptographically Secure Pseudo-Random Number Generator with high entropy (e.g., /dev/urandom)."),
            ("Key Encapsulation & Secure Storage", "Protect keys at rest using Hardware Security Modules (HSM), TPM chips, or secure key vaults (e.g. AWS KMS)."),
            ("Message Processing (Stream vs. Block)", "Data is ingested either continuous bit-by-bit (stream cipher) or partitioned into fixed chunks (block cipher, 128-bit)."),
            ("Authenticated Encryption (AEAD)", "Combine confidentiality with tamper detection in a single pass (e.g., AES-GCM or ChaCha20-Poly1305).")
        ],
        "Notes: Emphasize Kerckhoffs's principle. Security by obscurity always fails. The mathematical strength lies entirely in key entropy and algorithm design."
    )

    # --- SLIDE 5: Symmetric Key Encryption: AES Step-by-Step ---
    create_step_slide(
        "Symmetric Encryption: AES (Advanced Encryption Standard)",
        "Module 4: Symmetric Ciphers",
        [
            ("Architecture", "Symmetric cipher: identical shared secret key used for both encryption and decryption."),
            ("Block Size & Keys", "Operates on 128-bit blocks of data. Key sizes: 128-bit (10 rounds), 192-bit (12 rounds), 256-bit (14 rounds)."),
            ("High Performance", "Hardware-accelerated by modern CPUs (Intel AES-NI) processing multiple gigabytes per second."),
            ("Common Modes", "GCM (Galois/Counter Mode for authenticated encryption), CBC (Cipher Block Chaining with IV).")
        ],
        [
            ("Key Expansion (Rijndael Key Schedule)", "Master key is expanded into a sequence of round keys (e.g., 10 round keys for AES-128) using rotation and S-Box lookup."),
            ("Initial Round: AddRoundKey", "The 128-bit plaintext block is arranged in a 4x4 byte state matrix and XORed with Round Key 0."),
            ("SubBytes (Non-Linear Byte Substitution)", "Each byte in the 4x4 state matrix is substituted with an independent byte from the Rijndael S-Box, providing confusion."),
            ("ShiftRows (Row Permutation)", "Row 0 shifts 0 bytes, Row 1 shifts 1 byte left, Row 2 shifts 2 bytes, Row 3 shifts 3 bytes, providing diffusion across columns."),
            ("MixColumns (Column Matrix Multiplication)", "Each 4-byte column is transformed using matrix multiplication in Galois Field GF(2^8). (Omitted in final round)."),
            ("AddRoundKey & Final Round", "State matrix is XORed with the round key. After 10-14 rounds, the final state yields the scrambled 128-bit Ciphertext.")
        ],
        "Notes: Walk the audience through the four transformation stages of AES: SubBytes (confusion), ShiftRows & MixColumns (diffusion), and AddRoundKey (key insertion)."
    )

    # --- SLIDE 6: Asymmetric Encryption: RSA Algorithm Step-by-Step ---
    create_step_slide(
        "Asymmetric Encryption: RSA Mathematical Workflow",
        "Module 5: Public Key Cryptography",
        [
            ("Dual-Key Concept", "Public key (distributed openly for encryption) + Private key (kept secret for decryption)."),
            ("Mathematical Foundation", "Based on the computational intractability of factoring large semi-prime numbers (Integer Factorization Problem)."),
            ("Primary Use Cases", "Key encapsulation, secure key exchange, and digital signatures. Slower than symmetric ciphers."),
            ("Recommended Size", "Minimum 2048-bit; standard enterprise recommendation is 3072-bit or 4096-bit.")
        ],
        [
            ("Prime Number Selection", "Select two distinct, large random prime numbers p and q (each 1024 to 2048 bits in size)."),
            ("Compute Modulus & Totient", "Calculate modulus n = p * q (public component). Calculate Euler's totient phi(n) = (p - 1) * (q - 1)."),
            ("Select Public Exponent e", "Choose e such that 1 < e < phi(n) and gcd(e, phi(n)) = 1. Standard choice: e = 65537 (2^16 + 1)."),
            ("Derive Private Exponent d", "Calculate d using the Extended Euclidean Algorithm such that (d * e) mod phi(n) = 1 (modular inverse)."),
            ("Encryption Process: C = M^e mod n", "Sender takes plaintext message M (as integer < n) and computes ciphertext C using recipient's public key (e, n)."),
            ("Decryption Process: M = C^d mod n", "Recipient uses private key d to restore original message: M = C^d mod n = (M^e)^d mod n = M.")
        ],
        "Notes: Clarify that RSA is not used to encrypt big files directly due to speed. It encrypts an ephemeral symmetric key (Hybrid Cryptography), which then encrypts the bulk payload."
    )

    # --- SLIDE 7: Key Exchange: Diffie-Hellman Protocol ---
    create_step_slide(
        "Key Exchange: Diffie-Hellman Protocol (DH / ECDH)",
        "Module 6: Key Distribution",
        [
            ("The Fundamental Problem", "How do two strangers communicate secretly over an insecure channel without meeting beforehand?"),
            ("Core Innovation", "Invented by Whitfield Diffie & Martin Hellman (1976). Allows shared secret derivation over open networks."),
            ("Mathematical Trapdoor", "Relies on the Discrete Logarithm Problem: easy to compute g^a mod p, practically impossible to find a given g^a mod p."),
            ("Modern Standard", "Elliptic Curve Diffie-Hellman (ECDH) on Curve25519 or secp256r1 for faster computation & smaller keys.")
        ],
        [
            ("Public Parameter Agreement", "Alice and Bob openly agree on a large prime modulus p and base generator g (e.g. RFC 3526 groups)."),
            ("Private Secret Generation", "Alice chooses private integer 'a'. Bob chooses private integer 'b'. Neither reveals their private integer."),
            ("Public Value Computation", "Alice computes public key A = g^a mod p. Bob computes public key B = g^b mod p."),
            ("Exchange of Public Keys", "Alice sends A to Bob over the untrusted network. Bob sends B to Alice over the network."),
            ("Shared Secret Derivation", "Alice computes S = B^a mod p. Bob computes S = A^b mod p. Since (g^b)^a = (g^a)^b = g^(ab) mod p, both get identical S!"),
            ("Symmetric Session Key Generation", "Pass shared secret S through a Key Derivation Function (HKDF) to generate clean AES session keys.")
        ],
        "Notes: Highlight that an eavesdropper sees p, g, A, and B, but cannot determine S without computing the discrete logarithm, which is computationally infeasible."
    )

    # --- SLIDE 8: Cryptographic Hash Functions: SHA-256 & HMAC ---
    create_step_slide(
        "Cryptographic Hashing: SHA-256 & HMAC Verification",
        "Module 7: Integrity & Authentication",
        [
            ("One-Way Function", "Transforms arbitrary-length input into fixed-size digest (256 bits). Mathematically irreversible."),
            ("Avalanche Effect", "Flipping a single bit in the input radically alters more than 50% of the output hash digest."),
            ("Collision Resistance", "Infeasible to find two different inputs M1 and M2 such that Hash(M1) == Hash(M2)."),
            ("HMAC Extension", "Hash-based Message Authentication Code: HMAC = Hash((K ^ opad) || Hash((K ^ ipad) || M)).")
        ],
        [
            ("Pre-processing & Bit Padding", "Append bit '1' to message, followed by '0's, ending with a 64-bit integer representing original bit length (multiple of 512)."),
            ("Message Parsing", "Divide padded bitstream into N sequential 512-bit message blocks (M_1, M_2, ..., M_N)."),
            ("Initialize Working Variables", "Load 8 initial hash state registers (H0 through H7) with fractional square root constants of first 8 primes."),
            ("Message Schedule Expansion", "For each 512-bit block, expand 16 32-bit words into 64 words (W_0 to W_63) using bitwise rotations and shifts."),
            ("64-Round Compression Loop", "Execute 64 rounds of non-linear compression using Majority (Maj) and Choose (Ch) logic with round constants K_t."),
            ("Digest Output Concatenation", "Add compressed state back to H0..H7. Final 256-bit hexadecimal string serves as the unique digital fingerprint.")
        ],
        "Notes: Emphasize that hashing is not encryption—there is no key to decrypt it. It is an integrity check and the core building block for blockchain, signatures, and password verification."
    )

    # --- SLIDE 9: Digital Signatures & Public Key Infrastructure ---
    create_step_slide(
        "Digital Signatures: Non-Repudiation & PKI Trust Chain",
        "Module 8: Digital Trust & PKI",
        [
            ("Core Purpose", "Provides Authentication (who sent it), Integrity (unaltered content), and Non-Repudiation (cannot deny signing)."),
            ("Asymmetric Reversal", "Signing uses the Sender's Private Key; Verification uses the Sender's Public Key."),
            ("Algorithms", "RSA-PSS, ECDSA (Elliptic Curve Digital Signature Algorithm), Ed25519."),
            ("X.509 Certificates", "Digital identity passport binding a public key to an organization, signed by a trusted Certificate Authority (CA).")
        ],
        [
            ("Document Hashing (Signer)", "Sender generates cryptographic hash of document: H = SHA-256(Document). Ensures constant small payload."),
            ("Signature Generation (Signer)", "Sender encrypts hash H with their private key: Signature = Sign(H, PrivateKey_Sender)."),
            ("Bundle & Transmission", "Signer packages the original Document + Digital Signature + X.509 Digital Certificate and sends to recipient."),
            ("Certificate Chain Verification (Recipient)", "Recipient verifies signer's X.509 certificate up to root CA (e.g., DigiCert, Let's Encrypt) to confirm identity."),
            ("Signature Decryption (Recipient)", "Recipient decrypts signature using Signer's verified Public Key to extract original hash H_expected."),
            ("Integrity Comparison Check", "Recipient computes H_actual = SHA-256(Document). If H_actual == H_expected, document is authentic and untampered!")
        ],
        "Notes: Point out the distinction between confidentiality and non-repudiation. In signatures, everyone can verify with the public key, but only one entity could have signed it."
    )

    # --- SLIDE 10: Real-World Hybrid Crypto: TLS 1.3 Handshake ---
    create_step_slide(
        "Real-World Implementation: The TLS 1.3 Handshake Flow",
        "Module 9: Network Security",
        [
            ("Hybrid Cryptography", "Blends asymmetric crypto (for identity & key exchange) with symmetric crypto (for high-speed bulk data)."),
            ("TLS 1.3 Efficiency", "Completes handshake in 1-RTT (Round Trip Time) vs 2-RTT in TLS 1.2, reducing latency drastically."),
            ("Perfect Forward Secrecy", "Ephemeral ECDH keys ensure compromised server private keys cannot decrypt past recorded sessions."),
            ("Zero Plaintext Leakage", "All handshake metadata after ServerHello is fully encrypted.")
        ],
        [
            ("ClientHello (1-RTT Start)", "Client sends supported cipher suites (e.g. TLS_AES_256_GCM_SHA384) + Client Key Share (ephemeral ECDH public key)."),
            ("ServerHello & Key Selection", "Server picks cipher suite, responds with Server Key Share (its ephemeral ECDH public key). Both can now compute master secret!"),
            ("Encrypted Certificate Transmission", "Server sends encrypted X.509 Certificate and signs handshake transcript with private key for authentication."),
            ("Client Verification", "Client validates certificate against local root store, verifies server signature, and confirms identity."),
            ("Session Key Generation", "Both endpoints derive identical symmetric encryption keys (AES-256-GCM) and MAC keys using HKDF."),
            ("Encrypted Data Transmission", "Application HTTP traffic (HTTPS) is transmitted at wire speed using symmetric authenticated encryption.")
        ],
        "Notes: TLS 1.3 is the bedrock of internet security today. Show how Diffie-Hellman, RSA/ECDSA, SHA-256, and AES all harmonize into one secure connection."
    )

    # --- SLIDE 11: Enterprise Defense-in-Depth Architecture ---
    create_step_slide(
        "Enterprise Architecture: Multi-Layered Defense-in-Depth",
        "Module 10: Systemic Defense",
        [
            ("Core Doctrine", "Assume breach. Layer overlapping security controls so if any single defense fails, subsequent layers halt the attacker."),
            ("Zero Trust Philosophy", "'Never Trust, Always Verify'. Validate explicitly, use least privilege access, assume compromise."),
            ("SIEM & SOAR", "Centralized security information and event management coupled with automated incident response orchestration."),
            ("Human Element", "Phishing simulations and continuous security awareness training.")
        ],
        [
            ("Perimeter & Edge Layer", "Deploy Next-Generation Firewalls (NGFW), DDoS scrubbing (Cloudflare/Akamai), and Web Application Firewalls (WAF)."),
            ("Network Segmentation Layer", "Implement software-defined micro-segmentation, isolated VLANs, and strict zero-trust network access (ZTNA)."),
            ("Endpoint Protection Layer", "Equip servers and laptops with EDR/XDR agents, full-disk encryption (BitLocker/FileVault), and strict USB policies."),
            ("Application Security Layer", "Integrate SAST/DAST in CI/CD pipeline, API gateways with rate limiting, and strict input sanitization against OWASP Top 10."),
            ("Data & Identity Layer", "Enforce Multi-Factor Authentication (FIDO2/WebAuthn), Least Privilege RBAC, and AES-256 encryption at rest."),
            ("Resilience & Recovery Layer", "Maintain immutable air-gapped backups, disaster recovery sites, and automated incident response runbooks.")
        ],
        "Notes: Emphasize that cryptography protects the data layer, but without perimeter, identity, and endpoint controls, attackers simply steal keys or capture data in memory."
    )

    # --- SLIDE 12: Future Frontiers & Post-Quantum Cryptography ---
    create_step_slide(
        "Future Frontiers: Post-Quantum Cryptography & Zero Trust",
        "Module 11: Emerging Trends",
        [
            ("Quantum Computing Threat", "Shor's Algorithm will crack RSA and ECC. Grover's Algorithm cuts symmetric key strength in half (AES-128 becomes 64-bit)."),
            ("Harvest Now, Decrypt Later", "Nation-state adversaries currently capture encrypted traffic to decrypt once quantum computers mature."),
            ("NIST PQC Standards (2024)", "NIST finalized initial quantum-resistant standards: ML-KEM (Kyber), ML-DSA (Dilithium), SLH-DSA (SPHINCS+)."),
            ("Zero Knowledge Proofs (ZKP)", "Cryptographic proofs enabling one party to prove a statement is true without revealing underlying information.")
        ],
        [
            ("Cryptographic Inventory & Discovery", "Audit and scan all systems, certificates, and codebases to identify vulnerable classical algorithms (RSA/ECC)."),
            ("Crypto-Agility Architecture", "Refactor software to allow swapping cryptographic algorithms dynamically without rewriting application code."),
            ("Deploy Hybrid PQC Algorithms", "Implement hybrid key exchanges combining classical ECDH with post-quantum ML-KEM (as supported in modern Chrome & OpenSSH)."),
            ("Adopt Homomorphic Encryption", "Implement Fully Homomorphic Encryption (FHE) to allow analytics and AI training directly on encrypted datasets."),
            ("Zero-Knowledge Identity Verification", "Authenticate users and verify credentials without exposing passwords, PII, or biometric data across the network.")
        ],
        "Notes: Point out that transition to Post-Quantum Cryptography takes years. Organizations must build crypto-agility today before cryptanalytically relevant quantum computers arrive."
    )

    # --- SLIDE 13: Summary & Implementation Checklist ---
    create_step_slide(
        "Executive Summary: Security & Cryptography Best Practices",
        "Module 12: Actionable Checklist",
        [
            ("Key Takeaways", "Cybersecurity requires systemic defense across all layers; cryptography ensures mathematically backed trust & privacy."),
            ("Fail-Safe Design", "Design systems assuming components will fail or be compromised."),
            ("Compliance & Standards", "Align with ISO 27001, NIST CSF, PCI-DSS, and SOC 2 Type II frameworks."),
            ("Continuous Improvement", "Conduct red team vs blue team exercises and regular independent security audits.")
        ],
        [
            ("Never Roll Your Own Crypto", "Always leverage battle-tested standard libraries (OpenSSL, Libsodium, BoringSSL, WebCrypto API)."),
            ("Phase Out Legacy Algorithms", "Immediately deprecate MD5, SHA-1, DES, 3DES, RC4, and RSA keys under 2048 bits."),
            ("Enforce Strong Key Lifecycle", "Automate key generation, rotation (every 90 days), secure HSM storage, and immediate revocation protocols."),
            ("Universal MFA & Identity Hardening", "Eliminate passwords where possible; enforce phishing-resistant hardware security keys (FIDO2 / YubiKey)."),
            ("Immutable Backups & Incident Drills", "Store immutable offline backups and run quarterly tabletop incident response drills for ransomware readiness.")
        ],
        "Notes: Wrap up by emphasizing the Golden Rule: 'Never write your own cryptography algorithms.' Always rely on audited, standardized implementations and layered defense."
    )

    # Save presentation
    output_path = r"c:\Users\SAI CHARAN\Downloads\A1 boys hostel\A1 boys hostel\Cybersecurity_and_Cryptography_Methods.pptx"
    prs.save(output_path)
    print(f"Presentation successfully saved to: {output_path}")

if __name__ == "__main__":
    create_presentation()
