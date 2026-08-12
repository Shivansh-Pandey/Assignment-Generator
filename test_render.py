import os
import json
import app

# Real test payload matching user's exact failing screenshots
sample_data = {
    "title": "Magnetic Effects of Electric Current and Magnetism",
    "subject": "Physics",
    "class_level": "Class 12",
    "topic": "Magnetic Effects of Current and Magnetism",
    "difficulty": "ADVANCED (CBSE Board / Competitive Prep)",
    "total_marks": 30,
    "sections": [
        {
            "section_label": "Section A - Multiple Choice Questions",
            "instructions": "Choose the correct option. Each question carries 1 mark.",
            "questions": [
                {
                    "number": 1,
                    "marks": 1,
                    "type": "MCQ",
                    "question": "A loop carries a current I' in the same direction as I. What is the net force on the loop?",
                    "options": [
                        "A. (μ₀ × I × I' × b) / (2 × π × r)",
                        "B. (μ₀ × I × I' × b × a) / (2 × π × r × (r + a))",
                        "C. (μ₀ × I × I' × b × a) / (2 × π × r)",
                        "D. (μ₀ × I × I' × b × a) / (2 × π × (r + a))"
                    ],
                    "answer": "B"
                },
                {
                    "number": 2,
                    "marks": 1,
                    "type": "MCQ",
                    "question": "What is the magnetic field produced by a straight current-carrying wire of finite length?",
                    "options": [
                        "A. (μ0 * I) / (2 * pi * r)",
                        "B. (μ0 * I) / (4 * pi * r) * (sin(theta_1) + sin(theta_2))",
                        "C. (μ0 * I) / (4 * pi * r) * (cos(theta_1) + cos(theta_2))",
                        "D. (μ0 * I) / (2 * pi * r) * (sin(θ))"
                    ],
                    "answer": "B"
                }
            ]
        },
        {
            "section_label": "Section B - Short Answer Questions",
            "instructions": "Answer in 2-3 sentences. Each question carries 2 marks.",
            "questions": [
                {
                    "number": 3,
                    "marks": 2,
                    "type": "SHORT",
                    "question": "A proton, a deuteron, and an α particle, all having the same kinetic energy, enter a region of uniform magnetic field perpendicular to their velocities. Compare their radii of the circular paths. (Mass of deuteron = 2 × Mass of proton, Charge of deuteron = Charge of proton; Mass of α particle = 4 × Mass of proton, Charge of α particle = 2 × Charge of proton).",
                    "options": [],
                    "answer": ""
                },
                {
                    "number": 4,
                    "marks": 2,
                    "type": "SHORT",
                    "question": "A long straight wire carries a current of 30 A. An electron is moving with a velocity of 10^5 m/s parallel to the wire at a distance of 10 cm from it. Calculate the force experienced by the electron. (Given: μ₀ = 4 × π × 10^-7 T m/A, charge of electron = 1.6 × 10^-19 C).",
                    "options": [],
                    "answer": ""
                },
                {
                    "number": 5,
                    "marks": 2,
                    "type": "SHORT",
                    "question": "(a) State Ampere's Circuital Law. (1 mark)\n(b) Use it to find the magnetic field due to a long straight current-carrying wire. (2 marks)\n(c) Find the magnetic field for:\n(i) r < a\n(ii) a < r < b\n(iii) r > b",
                    "options": [],
                    "answer": ""
                }
            ]
        }
    ],
    "marking_scheme": [
        {
            "question_number": 1,
            "marks": 1,
            "type": "MCQ",
            "question_summary": "Net force on loop",
            "answer_key": "B",
            "mark_breakdown": ["Formula application - 1 mark"]
        }
    ]
}

print("Testing _prep pipeline on sample strings:")
print("Q3 Question:", app._prep(sample_data["sections"][1]["questions"][0]["question"]).encode('ascii', 'xmlcharrefreplace').decode())
print("Q4 Question:", app._prep(sample_data["sections"][1]["questions"][1]["question"]).encode('ascii', 'xmlcharrefreplace').decode())
print("Q1 Option A:", app._prep(sample_data["sections"][0]["questions"][0]["options"][0]).encode('ascii', 'xmlcharrefreplace').decode())

pdf_bytes = app.build_pdf(sample_data)
out_path = os.path.join(os.path.dirname(__file__), "sample_output.pdf")
with open(out_path, "wb") as f:
    f.write(pdf_bytes)

print(f"\nSUCCESS: Generated {out_path} ({len(pdf_bytes)} bytes)")
