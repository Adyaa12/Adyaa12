from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN

def add_title_slide(pres, title, subtitle):
    """Adds a title slide to the presentation."""
    slide_layout = pres.slide_layouts[0]
    slide = pres.slides.add_slide(slide_layout)
    title_shape = slide.shapes.title
    subtitle_shape = slide.placeholders[1]
    title_shape.text = title
    subtitle_shape.text = subtitle

def add_content_slide(pres, title, content_points):
    """Adds a content slide with a title and bullet points."""
    slide_layout = pres.slide_layouts[1]
    slide = pres.slides.add_slide(slide_layout)
    title_shape = slide.shapes.title
    body_shape = slide.placeholders[1]
    
    title_shape.text = title
    
    tf = body_shape.text_frame
    tf.clear()  # Clear existing text
    
    for point in content_points:
        p = tf.add_paragraph()
        p.text = point
        p.font.size = Pt(18)
        p.level = 0

# --- Main Script ---
prs = Presentation()
prs.slide_width = Inches(16)
prs.slide_height = Inches(9)

# Slide 1: Title Slide
add_title_slide(prs, 
    "A Day of Fun, Family & Celebration!", 
    "An Event Proposal for the [Client Company Name] Annual Family Day\nPrepared by: [Your Company Name]"
)

# Slide 2: Our Vision & Understanding
add_content_slide(prs, 
    "Our Vision & Understanding", 
    [
        "Our Goal: To design and execute a seamless, engaging, and memorable Family Day that celebrates your most valuable asset.",
        "Our Understanding: You want an event that shows appreciation, fosters community, and strengthens bonds.",
        "Our Promise: A stress-free, professionally managed event that allows you to relax and enjoy the day."
    ]
)

# Slide 3: The Theme: "Summer Fiesta Carnival"
add_content_slide(prs,
    "The Theme: 'Summer Fiesta Carnival'",
    [
        "Concept: A lively, interactive, and festive theme that appeals to all ages, from toddlers to grandparents.",
        "Why this theme?",
        "   \u2022 Universally Appealing: Everyone loves a carnival!",
        "   \u2022 Visually Stunning: Creates a vibrant, Instagram-worthy atmosphere.",
        "   \u2022 Highly Interactive: Encourages participation with games and activities.",
        "   \u2022 Scalable: Easily adaptable to your guest count and venue."
    ]
)

# Slide 4: Decor & Ambiance
add_content_slide(prs,
    "Decor & Ambiance: Bringing the Carnival to Life",
    [
        "Grand Entrance: A spectacular balloon arch in fiesta colors with a personalized welcome banner.",
        "Venue Transformation: Bright drapes, colorful bunting, and festive festoon lighting.",
        "Zonal Signage: Fun, carnival-style signs to guide guests to different activity zones.",
        "Photo Opportunities: A main photo booth with props and a roving, branded Instagram frame."
    ]
)

# Slide 5: Event Flow & Schedule
add_content_slide(prs,
    "Event Flow & Schedule",
    [
        "11:00 AM | Gates Open & Welcome (Music, Drinks, Registration)",
        "11:30 AM | Carnival Grounds Open (Games & Activities Begin)",
        "12:30 PM | The Fiesta Feast (Lunch Service)",
        "2:00 PM | Center Stage Spectacle (e.g., Magic Show)",
        "3:00 PM | Words from Leadership & Prize Giving",
        "3:30 PM | Sweet Treats & Mingle",
        "4:00 PM | Grand Finale & Farewell Gift Collection"
    ]
)

# Slide 6: Games & Activities (Zone 1: The Kids' Corner)
add_content_slide(prs,
    "Games & Activities (Zone 1: The Kids' Corner)",
    [
        "Toddler Zone (Ages 1-4): A safe, enclosed soft-play area, ball pit, and bubble machine.",
        "Kids' Creative Camp (Ages 5-12):",
        "   \u2022 Face Painting & Temporary Tattoos",
        "   \u2022 Balloon Twisting Artist",
        "   \u2022 Arts & Crafts Station (Decorate visors, paint tiles, etc.)"
    ]
)

# Slide 7: Games & Activities (Zone 2: Family Fun Fairway)
add_content_slide(prs,
    "Games & Activities (Zone 2: Family Fun Fairway)",
    [
        "Classic Carnival Stalls: Ring Toss, Knock Down Cans, Hook-a-Duck (with prizes!).",
        "Giant Garden Games: Giant Jenga, Giant Connect Four, Lawn Bowling.",
        "Team Challenges (Led by Emcee): Family Sack Race, Three-Legged Race, Tug-of-War."
    ]
)

# Slide 8: Entertainment & Live Acts
add_content_slide(prs,
    "Entertainment & Live Acts",
    [
        "Master of Ceremonies (MC): An energetic host to guide the event and engage the crowd.",
        "Live Band / DJ: Playing a family-friendly mix of popular and themed music.",
        "Main Stage Show: A 30-minute interactive Magic & Illusion Show.",
        "Roving Entertainers: Stilt Walker and a Caricature Artist for personal takeaways."
    ]
)

# Slide 9: Food & Beverages: A Culinary Carnival
add_content_slide(prs,
    "Food & Beverages: A Culinary Carnival",
    [
        "Concept: Interactive live food stalls to enhance the carnival experience.",
        "Savory Stalls: Taco & Burrito Bar, Gourmet Burger Stand, Live Pasta Station.",
        "Fun Food Carts: Popcorn Machine, Cotton Candy Stand, Ice Cream/Gelato Cart.",
        "Beverage Station: Fresh Lemonade, Iced Tea, Mocktail Bar, and water stations.",
        "(All dietary requirements will be catered for.)"
    ]
)

# Slide 10: Logistics & Flawless Execution
add_content_slide(prs,
    "Logistics & Flawless Execution",
    [
        "Event Management: A dedicated Event Manager and on-site team.",
        "Staffing: Professional and friendly hostesses, coordinators, and technical crew.",
        "Health & Safety: Designated First-Aid station with a certified medic.",
        "Registration: A smooth and efficient check-in process.",
        "Vendor Coordination: We handle all third-party vendors."
    ]
)

# Slide 11: Capturing the Memories
add_content_slide(prs,
    "Capturing the Memories",
    [
        "Professional Photography & Videography: To capture candid moments and create a highlight reel.",
        "Themed Photo Booth: With instant prints for guests to take home.",
        "Farewell Gift: A custom-branded tote bag for each family with treats and keepsakes."
    ]
)

# Slide 12: Why Choose Us?
add_content_slide(prs,
    "Why Choose [Your Company Name]?",
    [
        "End-to-End Solution: We handle everything from concept to cleanup.",
        "Proven Experience: A successful track record of executing large-scale events.",
        "Creative Excellence: We don't just plan events; we create experiences.",
        "Vendor Network: Strong relationships with the best vendors.",
        "Client-Centric Approach: Your vision is at the core of our planning."
    ]
)

# Slide 13: Next Steps & Q&A
add_content_slide(prs,
    "Next Steps & Q&A",
    [
        "Next Steps:",
        "   1. Discuss budget and finalize guest count.",
        "   2. Conduct a joint site visit.",
        "   3. Refine proposal based on feedback.",
        "   4. Contract signing and event kickoff!",
        "\nThank You!\n[Your Name/Company Name]\n[Your Phone Number] | [Your Email] | [Your Website]"
    ]
)

# Save the presentation
file_path = "Corporate_Family_Day_Proposal.pptx"
prs.save(file_path)

print(f"Presentation saved successfully as '{file_path}'")
