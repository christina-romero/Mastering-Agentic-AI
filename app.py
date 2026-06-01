"""Access Model Operating System.

An interactive, game-like training tool for Future 2 Guides to learn, practice,
and implement the Access Model. Built with Streamlit + pandas + openpyxl only.

The Excel workbook (access_model_training_data_v4.xlsx) is the only data source.
If it is missing, a representative one is generated from the embedded fallback so
the app always runs (e.g. on Streamlit Community Cloud).

Run the app:        streamlit run app.py
Build workbook only: python app.py build
"""

import io
import math
import os
import sys
import zlib

import pandas as pd
import streamlit as st

DATA_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)),
                         "access_model_training_data_v4.xlsx")

# ---------------------------------------------------------------------------
# Grade-band language rule. Never use the word "ages".
# ---------------------------------------------------------------------------
GRADE_LABELS = {
    "3-4": "Grades 3-4",
    "5-6": "Grades 5-6",
    "7-8": "Grades 7-8",
    "All": "All grades",
    "Adult": "Staff",
}


def grade_label(band):
    band = str(band).strip()
    return GRADE_LABELS.get(band, band)


# ---------------------------------------------------------------------------
# Per-category color map (drives left-border accents and badges).
# ---------------------------------------------------------------------------
CATEGORY_COLORS = {
    "Self-Regulation": "#6366f1",
    "Repair Protocols": "#ec4899",
    "Timeback": "#0ea5e9",
    "Check Charts": "#14b8a6",
    "Launch Facilitation": "#f59e0b",
    "Experiences": "#8b5cf6",
    "Teacher Coaching": "#ef4444",
    "Principal Concerns": "#d946ef",
    "Public School Compliance": "#3b82f6",
    "Multilingual Learners": "#10b981",
    "Special Education": "#f97316",
    "Attendance": "#64748b",
}
DEFAULT_ACCENT = "#6366f1"


def accent_for(category):
    return CATEGORY_COLORS.get(str(category), DEFAULT_ACCENT)


# ---------------------------------------------------------------------------
# Missions: each is a (Category, Topic) pair selecting one real scenario row.
# ---------------------------------------------------------------------------
MISSIONS = [
    {"level": 1, "title": "Understanding Autonomy",
     "category": "Principal Concerns", "topic": "Autonomy Abuse"},
    {"level": 1, "title": "Building Contribution",
     "category": "Launch Facilitation", "topic": "Community Building"},
    {"level": 1, "title": "Support, Don't Solve",
     "category": "Timeback", "topic": "Navigation"},
    {"level": 2, "title": "Repair Protocols",
     "category": "Repair Protocols", "topic": "Individual Harm"},
    {"level": 2, "title": "Timeback Coaching",
     "category": "Timeback", "topic": "Low XP"},
    {"level": 2, "title": "Productive Struggle",
     "category": "Timeback", "topic": "Avoidance"},
    {"level": 3, "title": "Teacher Coaching",
     "category": "Teacher Coaching", "topic": "Lecture Dependency"},
    {"level": 3, "title": "Parent Conversations",
     "category": "Principal Concerns", "topic": "Parent Pushback"},
    {"level": 3, "title": "Complex Implementation",
     "category": "Principal Concerns", "topic": "Large Group Dysregulation"},
]
TOTAL_MISSIONS = len(MISSIONS)
LEVELS = [1, 2, 3]


# ---------------------------------------------------------------------------
# Distractor generation. Two wrong answers per scenario model realistic,
# well-intentioned implementation MISTAKES: one "over-functioning" (doing too
# much / solving for the student) and one "under-functioning" (stepping back
# too far / ignoring). Never punitive, exclusionary, shaming, or compliance.
# Keyed by the exact principle name from the workbook.
# ---------------------------------------------------------------------------
PRINCIPLE_DISTRACTORS = {
    "Autonomy Is the Currency": (
        "Remove all of the student's autonomy indefinitely so the choice cannot come up again.",
        "Ignore it and let the student keep full autonomy with no follow-up.",
    ),
    "Capability Through Friction": (
        "Solve the hard part for the student so they can move past the struggle quickly.",
        "Leave the student to struggle alone with no coaching or check-in.",
    ),
    "Safety First": (
        "Take over completely and manage every part of the situation for the student.",
        "Wait and hope the situation settles on its own before doing anything.",
    ),
    "Support, Don't Solve": (
        "Do the task for the student so it gets completed correctly.",
        "Provide no support and tell the student to figure it out on their own.",
    ),
    "Belonging Through Contribution": (
        "Hand the student all the responsibility at once so they can prove they belong.",
        "Move on without the student and hope they choose to join in later.",
    ),
    "If It Doesn't Work, It Isn't Finished": (
        "Redo the work for the student so the final product meets the standard.",
        "Accept the work as finished even though it does not yet meet the standard.",
    ),
    "Knowledge Moves in All Directions": (
        "Position yourself as the only source of answers and direct every exchange.",
        "Step out of the exchange entirely and let it drift wherever it goes.",
    ),
}
GENERIC_DISTRACTORS = (
    "Take over and handle the whole situation for the student.",
    "Step back entirely and leave the student to manage it alone.",
)


def make_choices(row):
    """Return (correct, over_functioning, under_functioning) answer strings."""
    correct = str(row["Strong_Response"]).strip()
    principle = str(row["Access_Model_Principle"]).strip()
    over, under = PRINCIPLE_DISTRACTORS.get(principle, GENERIC_DISTRACTORS)
    return correct, over, under


# ---------------------------------------------------------------------------
# Audit guardrail. No generated answer choice may contain banned content.
# Patterns target punitive / exclusionary / shaming / compliance-first /
# power-struggle language and "send away to be ready" moves. They are chosen
# NOT to match the workbook's aligned Strong_Response wording.
# ---------------------------------------------------------------------------
BANNED_PATTERNS = [
    "punish", "punishment", "detention", "suspend", "suspension", "expel",
    "expulsion", "kick out", "kicked out", "send him home", "send her home",
    "send them home", "send the student home", "send the student away",
    "send away", "remove the student from the room",
    "remove the student from class", "remove the student from learning",
    "remove him from", "remove her from", "remove them from learning",
    "to the office", "shame", "embarrass", "humiliate", "make an example",
    "because i said", "do as you are told", "do as you're told", "obey",
    "comply or", "zero tolerance", "call security", "lose recess",
    "miss recess", "no recess", "write lines", "to be ready",
    "until he is ready", "until she is ready", "until they are ready",
]


def audit_text(text):
    """Return the list of banned patterns found in a single choice string."""
    low = str(text).lower()
    return [p for p in BANNED_PATTERNS if p in low]


def audit_scenarios(scen_df):
    """Audit every generated answer choice across all scenarios.

    Returns a list of (row_index, choice_text, [violations]).
    """
    violations = []
    for idx, row in scen_df.iterrows():
        for choice in make_choices(row):
            hits = audit_text(choice)
            if hits:
                violations.append((idx, choice, hits))
    return violations


# ---------------------------------------------------------------------------
# Stable, deterministic choice ordering. Stored per key in session state so it
# is identical across reruns.
# ---------------------------------------------------------------------------
def ordered_choices(key, choices):
    orders = st.session_state.setdefault("orders", {})
    if key not in orders:
        seed = zlib.crc32(key.encode("utf-8"))
        idx = list(range(len(choices)))
        # Deterministic Fisher-Yates using the crc32 seed.
        for i in range(len(idx) - 1, 0, -1):
            seed = (1103515245 * seed + 12345) & 0x7FFFFFFF
            j = seed % (i + 1)
            idx[i], idx[j] = idx[j], idx[i]
        orders[key] = [choices[k] for k in idx]
    return orders[key]


# ===========================================================================
# Workbook generation (embedded fallback / data source builder).
# ===========================================================================
SCENARIOS_CSV = """Category,Age_Band,Topic,Scenario,Access_Model_Principle,Strong_Response,Feedback,Traditional_Teacher_Move,Guide_Move,Alpha_Practice,Access_Alternative
Self-Regulation,3-4,Frustration,Student cries and refuses work after one mistake.,Capability Through Friction,Teach calming strategy during Launch and practice it daily.,Self-regulation should be explicitly taught.,Traditional response,Guide coaching response,Alpha implementation,Future 2 implementation
Self-Regulation,3-4,Transition,Student melts down during transition.,Safety First,Use visual routine and regulation check-in.,Predictable routines build regulation.,Traditional response,Guide coaching response,Alpha implementation,Future 2 implementation
Self-Regulation,3-4,Waiting,Student becomes dysregulated while waiting.,"Support, Don't Solve",Teach waiting tools and practice during Launch.,Skills must be taught before expected.,Traditional response,Guide coaching response,Alpha implementation,Future 2 implementation
Self-Regulation,5-6,Peer Conflict,Student yells at peer during disagreement.,Belonging Through Contribution,Facilitate repair conversation and teach regulation tool.,Conflict requires skill-building.,Traditional response,Guide coaching response,Alpha implementation,Future 2 implementation
Self-Regulation,5-6,Loss of Game,Student throws materials after losing.,Capability Through Friction,Practice emotional recovery strategies.,Challenge builds capability.,Traditional response,Guide coaching response,Alpha implementation,Future 2 implementation
Self-Regulation,5-6,Feedback,Student shuts down after feedback.,"Support, Don't Solve",Normalize feedback and create reflection routine.,Feedback is growth information.,Traditional response,Guide coaching response,Alpha implementation,Future 2 implementation
Self-Regulation,7-8,Public Embarrassment,Student reacts aggressively after correction.,Safety First,De-escalate privately and teach repair process.,Preserve dignity while maintaining accountability.,Traditional response,Guide coaching response,Alpha implementation,Future 2 implementation
Self-Regulation,7-8,Stress,Student reports overwhelming stress.,"Support, Don't Solve",Use action planning and regulation supports.,"Support effort, don't remove responsibility.",Traditional response,Guide coaching response,Alpha implementation,Future 2 implementation
Self-Regulation,7-8,Academic Failure,Student gives up after low score.,Capability Through Friction,Review growth plan and next steps.,Persistence is teachable.,Traditional response,Guide coaching response,Alpha implementation,Future 2 implementation
Repair Protocols,All,Individual Harm,Student insults a peer.,Belonging Through Contribution,Require individual repair before re-entry.,Individual harm requires individual repair.,Traditional response,Guide coaching response,Alpha implementation,Future 2 implementation
Repair Protocols,All,Group Harm,Student disrupts entire cohort.,Belonging Through Contribution,Require group repair before re-entry.,Group harm requires group repair.,Traditional response,Guide coaching response,Alpha implementation,Future 2 implementation
Repair Protocols,All,Property Damage,Student breaks classroom materials.,Belonging Through Contribution,Require restitution and repair plan.,Community resources require stewardship.,Traditional response,Guide coaching response,Alpha implementation,Future 2 implementation
Repair Protocols,All,Digital Harm,Student posts hurtful comment online.,Knowledge Moves in All Directions,Repair with impacted individuals and community.,Digital actions have real consequences.,Traditional response,Guide coaching response,Alpha implementation,Future 2 implementation
Repair Protocols,All,Exclusion,Student intentionally excludes peer.,Belonging Through Contribution,Restore belonging through repair action.,Community matters.,Traditional response,Guide coaching response,Alpha implementation,Future 2 implementation
Timeback,3-4,Navigation,Student cannot navigate lessons.,"Support, Don't Solve",Teach navigation explicitly.,"Teach barriers, don't bypass them.",Traditional response,Guide coaching response,Alpha implementation,Future 2 implementation
Timeback,3-4,Reading Support,Student cannot read directions.,"Support, Don't Solve",Use scaffolds and intervention.,Access before independence.,Traditional response,Guide coaching response,Alpha implementation,Future 2 implementation
Timeback,5-6,Low XP,Student consistently earns minimal XP.,Autonomy Is the Currency,Use data-driven coaching plan.,Data drives coaching.,Traditional response,Guide coaching response,Alpha implementation,Future 2 implementation
Timeback,5-6,Gaming XP,Student rushes lessons.,"If It Doesn't Work, It Isn't Finished",Prioritize mastery evidence.,Mastery over speed.,Traditional response,Guide coaching response,Alpha implementation,Future 2 implementation
Timeback,5-6,Avoidance,Student avoids difficult subjects.,Capability Through Friction,Coach through challenge.,Difficulty is not failure.,Traditional response,Guide coaching response,Alpha implementation,Future 2 implementation
Timeback,7-8,Data Review,Student ignores performance data.,Autonomy Is the Currency,Build weekly action plan.,Ownership requires awareness.,Traditional response,Guide coaching response,Alpha implementation,Future 2 implementation
Timeback,7-8,Plateau,Student growth stalls.,"Support, Don't Solve",Analyze data and adjust goals.,"Use evidence, not assumptions.",Traditional response,Guide coaching response,Alpha implementation,Future 2 implementation
Timeback,7-8,Mastery Failure,Student repeatedly fails mastery checks.,Capability Through Friction,Target prerequisite gaps.,Intervention should be precise.,Traditional response,Guide coaching response,Alpha implementation,Future 2 implementation
Check Charts,All,No Progress,Student never updates check chart.,Autonomy Is the Currency,Tie autonomy to evidence of progress.,Visible progress matters.,Traditional response,Guide coaching response,Alpha implementation,Future 2 implementation
Check Charts,All,Inflated Reporting,Student marks work complete without evidence.,"If It Doesn't Work, It Isn't Finished",Require proof of mastery.,Evidence matters.,Traditional response,Guide coaching response,Alpha implementation,Future 2 implementation
Check Charts,All,Peer Verification,Students approve each other inaccurately.,Knowledge Moves in All Directions,Teach quality verification.,Peer systems need standards.,Traditional response,Guide coaching response,Alpha implementation,Future 2 implementation
Check Charts,All,Reflection,Student cannot explain progress.,"Support, Don't Solve",Use reflection prompts.,Metacognition is a skill.,Traditional response,Guide coaching response,Alpha implementation,Future 2 implementation
Check Charts,All,Ownership,Student waits for adult updates.,Autonomy Is the Currency,Student updates chart independently.,Ownership should be visible.,Traditional response,Guide coaching response,Alpha implementation,Future 2 implementation
Launch Facilitation,3-4,Morning Readiness,Students arrive dysregulated.,Safety First,Embed regulation routine in Launch.,Launch sets conditions for success.,Traditional response,Guide coaching response,Alpha implementation,Future 2 implementation
Launch Facilitation,3-4,Attention,Students are distracted.,"Support, Don't Solve",Use movement and engagement.,Teach attention skills.,Traditional response,Guide coaching response,Alpha implementation,Future 2 implementation
Launch Facilitation,5-6,Goal Setting,Students cannot articulate goals.,Autonomy Is the Currency,Practice daily goal-setting.,Goals create ownership.,Traditional response,Guide coaching response,Alpha implementation,Future 2 implementation
Launch Facilitation,5-6,Community Building,Students exclude peers.,Belonging Through Contribution,Use contribution-based activities.,Belonging follows contribution.,Traditional response,Guide coaching response,Alpha implementation,Future 2 implementation
Launch Facilitation,7-8,Data Reflection,Students ignore performance trends.,Autonomy Is the Currency,Review data in Launch.,Reflection drives action.,Traditional response,Guide coaching response,Alpha implementation,Future 2 implementation
Launch Facilitation,7-8,Mindset,Students blame others for outcomes.,Capability Through Friction,Use ownership prompts.,Ownership is teachable.,Traditional response,Guide coaching response,Alpha implementation,Future 2 implementation
Experiences,5-6,Team Project,One student does all the work.,Belonging Through Contribution,Redistribute roles and accountability.,Contribution should be shared.,Traditional response,Guide coaching response,Alpha implementation,Future 2 implementation
Experiences,5-6,Conflict,Team disagreement halts progress.,Knowledge Moves in All Directions,Facilitate structured problem-solving.,Conflict can build capability.,Traditional response,Guide coaching response,Alpha implementation,Future 2 implementation
Experiences,7-8,Prototype Failure,Project fails during testing.,"If It Doesn't Work, It Isn't Finished",Require iteration.,Failure is feedback.,Traditional response,Guide coaching response,Alpha implementation,Future 2 implementation
Experiences,7-8,Expert Feedback,Students reject expert critique.,Capability Through Friction,Use revision cycle.,Growth requires feedback.,Traditional response,Guide coaching response,Alpha implementation,Future 2 implementation
Experiences,All,Presentation,Student refuses public sharing.,"Support, Don't Solve",Provide scaffolded participation.,Build confidence progressively.,Traditional response,Guide coaching response,Alpha implementation,Future 2 implementation
Teacher Coaching,Adult,Lecture Dependency,Teacher lectures for most of block.,"Support, Don't Solve",Model facilitation techniques.,Guides coach rather than lecture.,Traditional response,Guide coaching response,Alpha implementation,Future 2 implementation
Teacher Coaching,Adult,Rescuing,Teacher answers every question.,Capability Through Friction,Coach wait time and questioning.,Rescue reduces independence.,Traditional response,Guide coaching response,Alpha implementation,Future 2 implementation
Teacher Coaching,Adult,Data Blindness,Teacher ignores Timeback data.,Autonomy Is the Currency,Use data in coaching conversations.,Data should guide decisions.,Traditional response,Guide coaching response,Alpha implementation,Future 2 implementation
Teacher Coaching,Adult,Behavior Control,Teacher relies on compliance.,Belonging Through Contribution,Build contribution systems.,Culture beats control.,Traditional response,Guide coaching response,Alpha implementation,Future 2 implementation
Teacher Coaching,Adult,Low Expectations,Teacher lowers rigor.,Capability Through Friction,Maintain support and challenge.,"High support, high expectations.",Traditional response,Guide coaching response,Alpha implementation,Future 2 implementation
Principal Concerns,All,Desk Throwing,Student throws desk.,Safety First,Follow crisis plan and repair process.,Safety first.,Traditional response,Guide coaching response,Alpha implementation,Future 2 implementation
Principal Concerns,All,Self-Harm,Student hits self.,Safety First,Activate safety supports and regulation instruction.,Safety precedes learning.,Traditional response,Guide coaching response,Alpha implementation,Future 2 implementation
Principal Concerns,All,Autonomy Abuse,Student misuses autonomy.,Autonomy Is the Currency,Remove autonomy temporarily and create re-earning path.,Accountability matters.,Traditional response,Guide coaching response,Alpha implementation,Future 2 implementation
Principal Concerns,All,Reading Gap,Student cannot access curriculum independently.,"Support, Don't Solve",Provide access scaffolds.,Access before independence.,Traditional response,Guide coaching response,Alpha implementation,Future 2 implementation
Principal Concerns,All,Parent Pushback,Parent demands traditional instruction.,"Support, Don't Solve",Explain model and outcomes.,Transparency builds trust.,Traditional response,Guide coaching response,Alpha implementation,Future 2 implementation
Public School Compliance,All,IEP Accommodation,Student requires accommodation.,"Support, Don't Solve",Honor accommodation while preserving ownership.,Compliance and agency coexist.,Traditional response,Guide coaching response,Alpha implementation,Future 2 implementation
Public School Compliance,All,Testing Window,Student refuses assessment.,Capability Through Friction,Coach participation and document supports.,Requirements still matter.,Traditional response,Guide coaching response,Alpha implementation,Future 2 implementation
Public School Compliance,All,Documentation,Teacher fails to document intervention.,Autonomy Is the Currency,Use required systems consistently.,Systems protect students.,Traditional response,Guide coaching response,Alpha implementation,Future 2 implementation
Public School Compliance,All,Attendance Reporting,Attendance records incomplete.,Belonging Through Contribution,Follow procedures immediately.,Accuracy matters.,Traditional response,Guide coaching response,Alpha implementation,Future 2 implementation
Public School Compliance,All,Safety Drill,Students disengage during drill.,Safety First,Reinforce importance and expectations.,Safety is everyone's responsibility.,Traditional response,Guide coaching response,Alpha implementation,Future 2 implementation
Multilingual Learners,3-4,Language Access,Student cannot understand directions.,Knowledge Moves in All Directions,Use visuals and peer support.,Language should not block access.,Traditional response,Guide coaching response,Alpha implementation,Future 2 implementation
Multilingual Learners,5-6,Academic Vocabulary,Student struggles with key terms.,"Support, Don't Solve",Pre-teach vocabulary.,Language scaffolds increase access.,Traditional response,Guide coaching response,Alpha implementation,Future 2 implementation
Multilingual Learners,5-6,Peer Collaboration,Student avoids discussion.,Belonging Through Contribution,Structure partner interactions.,Participation builds confidence.,Traditional response,Guide coaching response,Alpha implementation,Future 2 implementation
Multilingual Learners,7-8,Written Responses,Student knows answer but cannot express it.,"Support, Don't Solve",Offer language supports.,"Measure learning, not just language.",Traditional response,Guide coaching response,Alpha implementation,Future 2 implementation
Multilingual Learners,7-8,Family Communication,Family needs translation.,Belonging Through Contribution,Provide accessible communication.,Partnership matters.,Traditional response,Guide coaching response,Alpha implementation,Future 2 implementation
Special Education,3-4,Attention Needs,Student struggles to sustain attention.,"Support, Don't Solve",Use accommodations and regulation tools.,Support should increase access.,Traditional response,Guide coaching response,Alpha implementation,Future 2 implementation
Special Education,3-4,Sensory Needs,Student overwhelmed by environment.,Safety First,Use sensory supports proactively.,Prevention beats reaction.,Traditional response,Guide coaching response,Alpha implementation,Future 2 implementation
Special Education,5-6,Executive Function,Student forgets tasks.,Autonomy Is the Currency,Teach planning systems.,Skills can be taught.,Traditional response,Guide coaching response,Alpha implementation,Future 2 implementation
Special Education,5-6,Processing Speed,Student needs additional time.,"Support, Don't Solve",Adjust pacing while preserving rigor.,Equity is not sameness.,Traditional response,Guide coaching response,Alpha implementation,Future 2 implementation
Special Education,7-8,Self-Advocacy,Student does not request supports.,Knowledge Moves in All Directions,Teach self-advocacy routines.,Independence requires awareness.,Traditional response,Guide coaching response,Alpha implementation,Future 2 implementation
Attendance,All,Chronic Absence,Student misses multiple days weekly.,Belonging Through Contribution,Create re-engagement plan.,Connection drives attendance.,Traditional response,Guide coaching response,Alpha implementation,Future 2 implementation
Attendance,All,Returning Student,Student returns after long absence.,"Support, Don't Solve",Prioritize reconnection and planning.,Belonging first.,Traditional response,Guide coaching response,Alpha implementation,Future 2 implementation
Attendance,All,Family Barriers,Attendance impacted by family challenges.,Belonging Through Contribution,Partner with family and support systems.,Relationships matter.,Traditional response,Guide coaching response,Alpha implementation,Future 2 implementation
Attendance,All,Avoidance,Student avoids specific subject.,Capability Through Friction,Address root cause and build confidence.,Avoidance often masks struggle.,Traditional response,Guide coaching response,Alpha implementation,Future 2 implementation
Attendance,All,Tardiness,Student consistently arrives late.,Autonomy Is the Currency,Build accountability and ownership plan.,Consistency matters.,Traditional response,Guide coaching response,Alpha implementation,Future 2 implementation
Launch Facilitation,5-6,Check-In,Students cannot identify emotions.,"Support, Don't Solve",Teach emotional vocabulary during Launch.,Self-awareness precedes regulation.,Traditional response,Guide coaching response,Alpha implementation,Future 2 implementation
Launch Facilitation,7-8,Community Norms,Students challenge norms.,Belonging Through Contribution,Co-create and revisit expectations.,Ownership increases buy-in.,Traditional response,Guide coaching response,Alpha implementation,Future 2 implementation
Experiences,7-8,Uneven Participation,One student disengages from team.,Belonging Through Contribution,Assign meaningful contribution role.,Contribution drives belonging.,Traditional response,Guide coaching response,Alpha implementation,Future 2 implementation
Teacher Coaching,Adult,Overhelping,Teacher completes work for students.,"Support, Don't Solve",Coach gradual release.,Ownership must remain with student.,Traditional response,Guide coaching response,Alpha implementation,Future 2 implementation
Principal Concerns,All,Large Group Dysregulation,Several students escalate simultaneously.,Safety First,Use practiced regulation systems from Launch.,Systems matter under pressure.,Traditional response,Guide coaching response,Alpha implementation,Future 2 implementation
Public School Compliance,All,504 Plan,Student needs documented support.,"Support, Don't Solve",Implement supports consistently.,Compliance protects access.,Traditional response,Guide coaching response,Alpha implementation,Future 2 implementation
Multilingual Learners,5-6,Translation Dependence,Student waits for translation.,"Support, Don't Solve",Use scaffolds that build independence.,Goal is increasing access.,Traditional response,Guide coaching response,Alpha implementation,Future 2 implementation
Special Education,7-8,Transition Anxiety,Student struggles with schedule changes.,Safety First,Pre-teach changes and regulation tools.,Predictability supports success.,Traditional response,Guide coaching response,Alpha implementation,Future 2 implementation
Attendance,All,Low Belonging,Student says school doesn't matter.,Belonging Through Contribution,Increase contribution opportunities.,Belonging improves engagement.,Traditional response,Guide coaching response,Alpha implementation,Future 2 implementation
Timeback,7-8,Subject Avoidance,Student skips hardest subject.,Capability Through Friction,Use coaching and action plan.,Challenge should not be avoided.,Traditional response,Guide coaching response,Alpha implementation,Future 2 implementation
"""

PRINCIPLES_ROWS = [
    {
        "Principle": "Autonomy Is the Currency",
        "Plain_Language_Definition": "Independence is something students earn and keep through evidence of responsibility.",
        "Guide_Look_For": "Students setting their own goals, tracking their own data, and owning their next steps.",
        "Common_Misinterpretation": "Treating autonomy as a reward to grant or revoke at will, or as freedom with no accountability.",
        "Strong_Implementation": "Tie autonomy to visible evidence and give students a clear path to earn it back when it slips.",
        "Related_App_Categories": "Timeback, Check Charts, Launch Facilitation, Principal Concerns",
    },
    {
        "Principle": "Capability Through Friction",
        "Plain_Language_Definition": "Productive struggle is where real capability is built, not a sign something is wrong.",
        "Guide_Look_For": "Students staying with a hard task, using strategies, and recovering after a setback.",
        "Common_Misinterpretation": "Believing any struggle is good, or rushing to remove all difficulty so students never feel stuck.",
        "Strong_Implementation": "Keep the challenge high and the support high; coach through the friction instead of erasing it.",
        "Related_App_Categories": "Self-Regulation, Timeback, Experiences, Teacher Coaching",
    },
    {
        "Principle": "Safety First",
        "Plain_Language_Definition": "Physical and emotional safety come before any learning task, every time.",
        "Guide_Look_For": "Calm bodies, predictable routines, and students who know what to do when dysregulated.",
        "Common_Misinterpretation": "Confusing safety with control, or using safety as a reason to lower expectations afterward.",
        "Strong_Implementation": "Use practiced regulation systems and crisis plans, then return to learning and repair.",
        "Related_App_Categories": "Self-Regulation, Launch Facilitation, Principal Concerns, Special Education",
    },
    {
        "Principle": "Support, Don't Solve",
        "Plain_Language_Definition": "Help students access the work without removing the thinking that belongs to them.",
        "Guide_Look_For": "Scaffolds, questions, and wait time that leave the student doing the cognitive work.",
        "Common_Misinterpretation": "Sliding into doing the task for the student, or withholding support and calling it independence.",
        "Strong_Implementation": "Offer the smallest scaffold that unlocks access, then fade it as the student takes over.",
        "Related_App_Categories": "Timeback, Multilingual Learners, Special Education, Check Charts",
    },
    {
        "Principle": "Belonging Through Contribution",
        "Plain_Language_Definition": "Students belong when they contribute something real to the community.",
        "Guide_Look_For": "Every student holding a meaningful role and repairing harm to rejoin the group.",
        "Common_Misinterpretation": "Trying to make students feel they belong without giving them a way to contribute.",
        "Strong_Implementation": "Design roles and repair paths so contribution is the route back to belonging.",
        "Related_App_Categories": "Repair Protocols, Launch Facilitation, Experiences, Attendance",
    },
    {
        "Principle": "If It Doesn't Work, It Isn't Finished",
        "Plain_Language_Definition": "Work is complete only when it actually meets the standard, not when time runs out.",
        "Guide_Look_For": "Iteration, evidence of mastery, and students revising until the work holds up.",
        "Common_Misinterpretation": "Accepting effort or completion as mastery, or redoing the work for the student to finish it.",
        "Strong_Implementation": "Require proof of mastery and build in iteration cycles before work is called done.",
        "Related_App_Categories": "Check Charts, Timeback, Experiences",
    },
    {
        "Principle": "Knowledge Moves in All Directions",
        "Plain_Language_Definition": "Learning flows between students, guides, families, and the wider world, not just top-down.",
        "Guide_Look_For": "Peer teaching, student-led exchanges, and guides learning alongside students.",
        "Common_Misinterpretation": "Positioning the adult as the only source of knowledge, or letting peer exchange run with no standards.",
        "Strong_Implementation": "Build structures for peer teaching and feedback with clear quality expectations.",
        "Related_App_Categories": "Check Charts, Experiences, Multilingual Learners, Repair Protocols",
    },
]

LAUNCH_ROWS = [
    {"Age_Band": "3-4", "Lesson_Title": "Naming My Feelings",
     "Self_Regulation_or_Culture_Skill": "Emotional awareness", "Duration": "10 min",
     "Lesson_Description": "Students learn a small set of feeling words and point to how they feel today.",
     "Mastery_Evidence": "Student names their current feeling without prompting.",
     "Aligned_Categories": "Self-Regulation, Launch Facilitation"},
    {"Age_Band": "3-4", "Lesson_Title": "My Calm-Down Plan",
     "Self_Regulation_or_Culture_Skill": "Regulation strategy", "Duration": "15 min",
     "Lesson_Description": "Students practice one calming strategy and rehearse using it after a mistake.",
     "Mastery_Evidence": "Student demonstrates the strategy during a practiced trigger.",
     "Aligned_Categories": "Self-Regulation"},
    {"Age_Band": "5-6", "Lesson_Title": "Setting a Daily Goal",
     "Self_Regulation_or_Culture_Skill": "Goal setting and ownership", "Duration": "12 min",
     "Lesson_Description": "Students write one measurable goal for the day and name how they will track it.",
     "Mastery_Evidence": "Student states a specific, trackable goal.",
     "Aligned_Categories": "Launch Facilitation, Timeback, Check Charts"},
    {"Age_Band": "5-6", "Lesson_Title": "Contribution Circle",
     "Self_Regulation_or_Culture_Skill": "Community and belonging", "Duration": "15 min",
     "Lesson_Description": "Each student names one way they will contribute to the cohort today.",
     "Mastery_Evidence": "Student identifies a concrete contribution and follows through.",
     "Aligned_Categories": "Launch Facilitation, Experiences, Attendance"},
    {"Age_Band": "5-6", "Lesson_Title": "Reading My Emotions",
     "Self_Regulation_or_Culture_Skill": "Self-awareness", "Duration": "10 min",
     "Lesson_Description": "Students expand their emotional vocabulary and check in on intensity.",
     "Mastery_Evidence": "Student labels an emotion and its intensity.",
     "Aligned_Categories": "Self-Regulation, Launch Facilitation"},
    {"Age_Band": "7-8", "Lesson_Title": "Reading My Data",
     "Self_Regulation_or_Culture_Skill": "Data reflection and ownership", "Duration": "15 min",
     "Lesson_Description": "Students review weekly performance trends and choose one action for the week.",
     "Mastery_Evidence": "Student names a trend and a specific next action.",
     "Aligned_Categories": "Timeback, Launch Facilitation, Check Charts"},
    {"Age_Band": "7-8", "Lesson_Title": "Owning My Outcomes",
     "Self_Regulation_or_Culture_Skill": "Ownership mindset", "Duration": "12 min",
     "Lesson_Description": "Students reframe a setback as something within their control to improve.",
     "Mastery_Evidence": "Student states an internal, actionable next step.",
     "Aligned_Categories": "Launch Facilitation, Self-Regulation, Teacher Coaching"},
    {"Age_Band": "All", "Lesson_Title": "Repair and Re-Entry",
     "Self_Regulation_or_Culture_Skill": "Repair and accountability", "Duration": "15 min",
     "Lesson_Description": "Students rehearse the steps to repair harm and rejoin the community.",
     "Mastery_Evidence": "Student can describe the repair steps in order.",
     "Aligned_Categories": "Repair Protocols, Launch Facilitation"},
]

CHECK_CHART_ROWS = [
    {"Skill_Area": "Self-Regulation", "Check_Name": "Calm-Down Strategy",
     "Check_Description": "Student uses a regulation strategy independently after a setback.",
     "Evidence_Required": "Observed use of the strategy without adult prompting.",
     "Mastery_Criteria": "Demonstrated on three separate occasions.",
     "Aligned_Categories": "Self-Regulation, Launch Facilitation"},
    {"Skill_Area": "Goal Ownership", "Check_Name": "Daily Goal Set",
     "Check_Description": "Student sets and records a measurable daily goal.",
     "Evidence_Required": "Goal recorded in the student's own plan.",
     "Mastery_Criteria": "Consistent across one week.",
     "Aligned_Categories": "Launch Facilitation, Timeback"},
    {"Skill_Area": "Mastery Evidence", "Check_Name": "Proof of Mastery",
     "Check_Description": "Student shows work that meets the standard before marking it done.",
     "Evidence_Required": "Artifact or assessment meeting the criteria.",
     "Mastery_Criteria": "Independent mastery evidence, not completion alone.",
     "Aligned_Categories": "Check Charts, Timeback"},
    {"Skill_Area": "Reflection", "Check_Name": "Explain My Progress",
     "Check_Description": "Student explains what they learned and what is next.",
     "Evidence_Required": "Verbal or written reflection.",
     "Mastery_Criteria": "Accurate self-assessment tied to evidence.",
     "Aligned_Categories": "Check Charts, Launch Facilitation"},
    {"Skill_Area": "Peer Verification", "Check_Name": "Quality Check",
     "Check_Description": "Student verifies a peer's work against a shared standard.",
     "Evidence_Required": "Verification notes referencing the criteria.",
     "Mastery_Criteria": "Accurate, standards-based verification.",
     "Aligned_Categories": "Check Charts, Knowledge Moves in All Directions"},
    {"Skill_Area": "Chart Ownership", "Check_Name": "Update My Chart",
     "Check_Description": "Student updates their own check chart from evidence.",
     "Evidence_Required": "Student-initiated chart updates.",
     "Mastery_Criteria": "Independent updates over time.",
     "Aligned_Categories": "Check Charts, Autonomy"},
]

ALPHA_ROWS = [
    {"Alpha_Practice": "Whole-group lecture to deliver content",
     "Why_It_Works": "Delivers information to many students at once and feels efficient.",
     "Future_2_Constraint": "Students learn core content through mastery-based lessons, not lectures.",
     "Access_Aligned_Alternative": "Guide facilitates short workshops and coaches small groups based on data."},
    {"Alpha_Practice": "Teacher answers every student question immediately",
     "Why_It_Works": "Reduces friction and keeps students moving quickly.",
     "Future_2_Constraint": "Capability is built through productive struggle.",
     "Access_Aligned_Alternative": "Guide uses wait time and questioning so students do the thinking."},
    {"Alpha_Practice": "Compliance-based behavior charts",
     "Why_It_Works": "Creates short-term order and clear expectations.",
     "Future_2_Constraint": "Belonging and contribution drive culture, not compliance.",
     "Access_Aligned_Alternative": "Guide builds contribution roles and repair routines."},
    {"Alpha_Practice": "Grading for completion",
     "Why_It_Works": "Rewards effort and is simple to track.",
     "Future_2_Constraint": "Work is finished only when it meets the standard.",
     "Access_Aligned_Alternative": "Guide requires mastery evidence and iteration before work is done."},
    {"Alpha_Practice": "Adult-managed progress tracking",
     "Why_It_Works": "Keeps records accurate and centralized.",
     "Future_2_Constraint": "Autonomy is earned through student-owned evidence.",
     "Access_Aligned_Alternative": "Students own and update their check charts from evidence."},
    {"Alpha_Practice": "Removing challenge for struggling students",
     "Why_It_Works": "Reduces frustration in the moment.",
     "Future_2_Constraint": "High support pairs with high expectations.",
     "Access_Aligned_Alternative": "Guide keeps rigor and adds targeted scaffolds and coaching."},
]

BRAINLIFT_ROWS = [
    {"Reference_Topic": "The Access Model overview",
     "Reference_Summary": "Education that removes barriers to access while keeping ownership and rigor with the student.",
     "Detailed_Context": "The Access Model starts from a simple distinction: a barrier is anything that "
        "blocks a student from reaching learning they are capable of, while the learning itself is the "
        "work that must stay with the student. A Guide's job is to remove or scaffold the barrier "
        "without removing the thinking. That is why access and rigor are not in tension here. Lowering "
        "the standard is not access; it is a different, smaller goal. Real access means the student "
        "reaches the same high bar with the supports they need, and then carries more of the load over "
        "time as the supports fade.",
     "Guide_Application": "Before acting, name the specific barrier (reading, navigation, regulation, "
        "language, executive function). Then choose the smallest support that unlocks access and plan "
        "how it fades, rather than defaulting to doing the task or dropping the expectation.",
     "Source_Note": "Future 2 foundational framing."},
    {"Reference_Topic": "Guides, not teachers",
     "Reference_Summary": "Adults coach, facilitate, and remove barriers rather than lecture or rescue.",
     "Detailed_Context": "In a traditional model the adult is the primary source of content and the "
        "engine of the room: they lecture, they answer, they manage. In the Access Model the platform "
        "and mastery-based lessons carry core content, which frees the adult to do what software cannot: "
        "coach motivation, build culture, read data, and protect productive struggle. The shift is from "
        "delivering information to engineering the conditions in which students learn and own their "
        "growth. Lecturing and rescuing both signal that the adult, not the student, owns the work.",
     "Guide_Application": "Measure your block by how much thinking students did, not how much you "
        "talked. Use wait time and questions, coach in small groups from data, and resist answering "
        "what a student can reach with one good prompt.",
     "Source_Note": "Future 2 role definition."},
    {"Reference_Topic": "Launch as the daily foundation",
     "Reference_Summary": "A short daily block that builds regulation, culture, goals, and reflection skills.",
     "Detailed_Context": "Launch is the part of the day where the non-academic skills that make "
        "everything else possible are explicitly taught and practiced: emotional regulation, goal "
        "setting, reflection, and community. The premise is that these are skills, not traits, and that "
        "they must be taught before they are expected. Conditions for a strong learning block are set "
        "in Launch, so when dysregulation, avoidance, or conflict appear later, the response is to "
        "practice the skill in Launch, not to react in the moment.",
     "Guide_Application": "When a recurring behavior shows up, trace it back to a Launch skill and "
        "build daily practice for it. Treat Launch as your highest-leverage prevention time, not a "
        "warm-up.",
     "Source_Note": "Launch facilitation guidance."},
    {"Reference_Topic": "Timeback and mastery",
     "Reference_Summary": "Mastery-based learning where evidence, not speed, defines completion.",
     "Detailed_Context": "Timeback is the learning platform, and its core rule is that mastery is "
        "defined by evidence, not by time spent or lessons clicked through. Speed and XP can be gamed; "
        "mastery cannot. This is why rushing, inflated reporting, and ignoring data are treated as "
        "coaching opportunities rather than discipline problems: they are signals about ownership and "
        "understanding. Data is the shared, objective language a Guide and student use to decide what "
        "happens next, which keeps coaching grounded in evidence instead of assumptions.",
     "Guide_Application": "Coach from the data weekly: name a trend, agree on one action, and require "
        "mastery evidence before work is called done. Prioritize precise intervention on prerequisite "
        "gaps over pushing pace.",
     "Source_Note": "Timeback platform principles."},
    {"Reference_Topic": "Repair over punishment",
     "Reference_Summary": "Harm is addressed through repair and re-entry, restoring belonging through contribution.",
     "Detailed_Context": "When harm happens, the Access Model asks what repair is owed and how the "
        "student rejoins the community, rather than what punishment fits. The logic is matched repair: "
        "individual harm calls for individual repair, group harm for group repair, and property harm "
        "for restitution, each completed before re-entry. Belonging is the lever, not exclusion, "
        "because students are most likely to change when contribution is the path back in. Repair holds "
        "students fully accountable while keeping them inside the community that helps them grow.",
     "Guide_Application": "Match the repair to the harm and make completing it the route back to the "
        "group. Avoid responses that remove a student from learning unless there is an immediate safety "
        "concern.",
     "Source_Note": "Repair protocol guidance."},
    {"Reference_Topic": "Productive struggle",
     "Reference_Summary": "Friction is where capability is built; support stays high while challenge stays high.",
     "Detailed_Context": "Capability is built in the zone where a task is hard but reachable with "
        "effort and strategy. Friction is therefore a feature, not a malfunction, and removing all "
        "difficulty quietly removes the learning. The Access Model rejects the trade-off between support "
        "and expectations: the strong move is high support and high expectations together. Two failure "
        "modes sit on either side, rescuing (high support, low expectations) and pressure (low support, "
        "high expectations), and the Guide's craft is staying in the high-high quadrant by coaching "
        "through the struggle instead of erasing it or abandoning the student to it.",
     "Guide_Application": "When a student is stuck, hold the bar and add a strategy or scaffold rather "
        "than the answer. Normalize struggle out loud and coach recovery after setbacks.",
     "Source_Note": "Capability through friction research framing."},
]


def build_workbook(path):
    """Generate a representative workbook matching the required schema."""
    scen = pd.read_csv(io.StringIO(SCENARIOS_CSV))
    principles = pd.DataFrame(PRINCIPLES_ROWS)
    launch = pd.DataFrame(LAUNCH_ROWS)
    checks = pd.DataFrame(CHECK_CHART_ROWS)
    alpha = pd.DataFrame(ALPHA_ROWS)
    brainlift = pd.DataFrame(BRAINLIFT_ROWS)

    sheets = {
        "Scenarios": scen,
        "Access Model Principles": principles,
        "Launch Lessons": launch,
        "Check Chart Alignments": checks,
        "Alpha to Access": alpha,
        "Brainlift References": brainlift,
    }

    # Data Dictionary derived from the other sheets.
    dd_rows = []
    dd_defs = {
        ("Scenarios", "Category"): "The Access Model topic area the scenario belongs to.",
        ("Scenarios", "Age_Band"): "Grade band the scenario applies to (3-4, 5-6, 7-8, All, or Staff).",
        ("Scenarios", "Topic"): "The specific situation within the category.",
        ("Scenarios", "Scenario"): "The classroom situation a Guide may face.",
        ("Scenarios", "Access_Model_Principle"): "The Access Model principle most relevant to the scenario.",
        ("Scenarios", "Strong_Response"): "The aligned Guide response (the correct training answer).",
        ("Scenarios", "Feedback"): "Why the strong response is aligned to the Access Model.",
        ("Scenarios", "Traditional_Teacher_Move"): "Placeholder filler column; not displayed in the app.",
        ("Scenarios", "Guide_Move"): "Placeholder filler column; not displayed in the app.",
        ("Scenarios", "Alpha_Practice"): "Placeholder filler column; not displayed in the app.",
        ("Scenarios", "Access_Alternative"): "Placeholder filler column; not displayed in the app.",
        ("Access Model Principles", "Principle"): "Name of the Access Model principle.",
        ("Access Model Principles", "Plain_Language_Definition"): "One-line definition of the principle.",
        ("Access Model Principles", "Guide_Look_For"): "What a Guide looks for when the principle is working.",
        ("Access Model Principles", "Common_Misinterpretation"): "The common trap or misreading of the principle.",
        ("Access Model Principles", "Strong_Implementation"): "What strong implementation looks like in practice.",
        ("Access Model Principles", "Related_App_Categories"): "Categories most connected to this principle.",
        ("Launch Lessons", "Age_Band"): "Grade band for the lesson.",
        ("Launch Lessons", "Lesson_Title"): "Title of the Launch lesson.",
        ("Launch Lessons", "Self_Regulation_or_Culture_Skill"): "The regulation or culture skill the lesson builds.",
        ("Launch Lessons", "Duration"): "Approximate lesson length.",
        ("Launch Lessons", "Lesson_Description"): "What happens during the lesson.",
        ("Launch Lessons", "Mastery_Evidence"): "Evidence that the skill was learned.",
        ("Launch Lessons", "Aligned_Categories"): "App categories the lesson supports.",
        ("Check Chart Alignments", "Skill_Area"): "The skill area the check measures.",
        ("Check Chart Alignments", "Check_Name"): "Name of the check.",
        ("Check Chart Alignments", "Check_Description"): "What the student must demonstrate.",
        ("Check Chart Alignments", "Evidence_Required"): "Evidence needed to mark the check.",
        ("Check Chart Alignments", "Mastery_Criteria"): "The bar for mastery of the check.",
        ("Check Chart Alignments", "Aligned_Categories"): "App categories the check supports.",
        ("Alpha to Access", "Alpha_Practice"): "A common practice from an Alpha/traditional setting.",
        ("Alpha to Access", "Why_It_Works"): "The partial truth that makes the practice tempting.",
        ("Alpha to Access", "Future_2_Constraint"): "The Future 2 constraint the practice conflicts with.",
        ("Alpha to Access", "Access_Aligned_Alternative"): "The Access-aligned alternative a Guide uses.",
        ("Brainlift References", "Reference_Topic"): "Topic of the reference.",
        ("Brainlift References", "Reference_Summary"): "Short summary of the reference.",
        ("Brainlift References", "Detailed_Context"): "Deeper Access Model context that explains the topic.",
        ("Brainlift References", "Guide_Application"): "How a Guide applies the topic in practice.",
        ("Brainlift References", "Source_Note"): "Where the reference comes from.",
    }
    for sheet_name, df in sheets.items():
        for col in df.columns:
            definition = dd_defs.get((sheet_name, col), "See sheet for details.")
            dd_rows.append({"Sheet": sheet_name, "Column": col, "Definition": definition})
    sheets["Data Dictionary"] = pd.DataFrame(dd_rows)

    with pd.ExcelWriter(path, engine="openpyxl") as writer:
        for name, df in sheets.items():
            df.to_excel(writer, sheet_name=name, index=False)
    return path


def ensure_workbook(path):
    if not os.path.exists(path):
        build_workbook(path)
    return path


# Build-only entry point (python app.py build) -- runs before any Streamlit code.
if __name__ == "__main__" and len(sys.argv) > 1 and sys.argv[1] == "build":
    build_workbook(DATA_PATH)
    print("Built workbook at", DATA_PATH)
    sys.exit(0)


# ===========================================================================
# Data loading
# ===========================================================================
@st.cache_data(show_spinner=False)
def load_data(path):
    ensure_workbook(path)
    return pd.read_excel(path, sheet_name=None, engine="openpyxl")


# ===========================================================================
# Streamlit app
# ===========================================================================
st.set_page_config(page_title="Access Model Operating System",
                   page_icon=None, layout="wide",
                   initial_sidebar_state="expanded")

CSS = """
<style>
:root {
  --am-bg1: #eef2ff;
  --am-bg2: #f8fafc;
  --am-ink: #0f172a;
  --am-muted: #64748b;
  --am-line: #e2e8f0;
  --am-accent: #6366f1;
}
html, body, [class*="css"] {
  font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, Helvetica,
    Arial, sans-serif;
}
.stApp {
  background: linear-gradient(160deg, var(--am-bg1) 0%, var(--am-bg2) 55%, #ffffff 100%);
}
.block-container { padding-top: 2.2rem; padding-bottom: 4rem; max-width: 1180px; }

/* Sidebar (dark) */
section[data-testid="stSidebar"] {
  background: #0f172a;
  border-right: 1px solid #1e293b;
}
section[data-testid="stSidebar"] * { color: #e2e8f0; }
.am-brand {
  padding: 6px 4px 14px 4px;
  border-bottom: 1px solid #1e293b;
  margin-bottom: 14px;
}
.am-brand .am-brand-main {
  font-size: 0.95rem; font-weight: 800; letter-spacing: 0.14em; color: #ffffff;
}
.am-brand .am-brand-sub {
  font-size: 0.72rem; letter-spacing: 0.18em; color: #94a3b8; text-transform: uppercase;
}
.am-navgroup {
  font-size: 0.66rem; font-weight: 700; letter-spacing: 0.16em; color: #64748b;
  text-transform: uppercase; margin: 16px 4px 6px 4px;
}
section[data-testid="stSidebar"] .stButton > button {
  width: 100%; text-align: left; justify-content: flex-start;
  border-radius: 10px; font-weight: 600; border: 1px solid transparent;
  padding: 0.5rem 0.75rem; margin-bottom: 2px;
}
section[data-testid="stSidebar"] .stButton > button[kind="secondary"] {
  background: transparent; color: #cbd5e1; border-color: transparent;
}
section[data-testid="stSidebar"] .stButton > button[kind="secondary"]:hover {
  background: #1e293b; color: #ffffff;
}
section[data-testid="stSidebar"] .stButton > button[kind="primary"] {
  background: linear-gradient(135deg, #6366f1, #8b5cf6); color: #ffffff;
  border-color: #6366f1; box-shadow: 0 6px 16px rgba(99,102,241,0.35);
}

/* Hero */
.am-hero {
  background: linear-gradient(120deg, #4f46e5 0%, #7c3aed 50%, #db2777 100%);
  color: #ffffff; border-radius: 22px; padding: 30px 34px;
  box-shadow: 0 18px 40px rgba(79,70,229,0.30); margin-bottom: 22px;
}
.am-hero h1 { color: #ffffff; font-size: 2.0rem; margin: 0 0 6px 0; font-weight: 800; }
.am-hero p { color: #e9d5ff; font-size: 1.02rem; margin: 0; }

/* Cards */
.am-card {
  background: #ffffff; border-radius: 16px; padding: 20px 22px;
  border: 1px solid var(--am-line); border-left: 6px solid var(--am-accent);
  box-shadow: 0 6px 18px rgba(15,23,42,0.06);
  transition: transform .15s ease, box-shadow .15s ease; margin-bottom: 14px;
}
.am-card:hover { transform: translateY(-3px); box-shadow: 0 14px 30px rgba(15,23,42,0.12); }
.am-card-title { font-size: 1.08rem; font-weight: 750; color: var(--am-ink); margin-bottom: 4px; }
.am-card-summary { font-size: 0.95rem; color: var(--am-muted); line-height: 1.5; }
.am-card-body { font-size: 0.94rem; color: #334155; line-height: 1.55; margin-top: 10px; }
.am-kv { margin-top: 8px; font-size: 0.92rem; color: #334155; }
.am-kv b { color: var(--am-ink); }

/* Stat cards */
.am-stat {
  background: #ffffff; border-radius: 16px; padding: 18px 20px;
  border: 1px solid var(--am-line); box-shadow: 0 6px 18px rgba(15,23,42,0.06);
  transition: transform .15s ease, box-shadow .15s ease; height: 100%;
}
.am-stat:hover { transform: translateY(-3px); box-shadow: 0 14px 30px rgba(15,23,42,0.12); }
.am-stat .am-stat-label {
  font-size: 0.72rem; letter-spacing: 0.10em; text-transform: uppercase;
  color: var(--am-muted); font-weight: 700;
}
.am-stat .am-stat-value { font-size: 1.9rem; font-weight: 800; color: var(--am-ink); margin-top: 4px; }
.am-stat .am-stat-sub { font-size: 0.82rem; color: var(--am-muted); margin-top: 2px; }

/* Badges & chips */
.am-badge {
  display: inline-block; padding: 3px 10px; border-radius: 999px;
  font-size: 0.74rem; font-weight: 700; letter-spacing: 0.03em;
}
.am-badge.done { background: #dcfce7; color: #166534; }
.am-badge.progress { background: #dbeafe; color: #1e40af; }
.am-badge.locked { background: #f1f5f9; color: #64748b; }
.am-badge.good { background: #dcfce7; color: #166534; }
.am-badge.revise { background: #fee2e2; color: #991b1b; }
.am-chip {
  display: inline-block; padding: 3px 10px; border-radius: 999px;
  background: #eef2ff; color: #4338ca; font-size: 0.74rem; font-weight: 700;
  margin-right: 6px;
}
.am-section { font-size: 1.3rem; font-weight: 800; color: var(--am-ink); margin: 6px 0 10px 0; }
.am-sub { color: var(--am-muted); margin-bottom: 16px; }

/* Graded answer rows */
.am-opt {
  border-radius: 12px; padding: 12px 16px; margin-bottom: 8px; font-size: 0.96rem;
  border: 1px solid var(--am-line); background: #ffffff;
}
.am-opt.correct { background: #ecfdf5; border-color: #6ee7b7; color: #065f46; font-weight: 600; }
.am-opt.wrong { background: #fef2f2; border-color: #fca5a5; color: #991b1b; font-weight: 600; }
.am-opt.muted { background: #f8fafc; color: #475569; }

/* Icons in card headers */
.am-card-head { display: flex; align-items: center; gap: 10px; margin-bottom: 4px; }
.am-card-head .am-card-title { margin-bottom: 0; }
.am-icon { display: inline-flex; flex: 0 0 auto; line-height: 0; }

/* Donut progress ring */
.am-donut-wrap { text-align: center; }
.am-donut-sub { font-size: 0.82rem; color: var(--am-muted); font-weight: 600; margin-top: 2px; }

/* Support x Expectations matrix */
.am-matrix-card {
  background: #ffffff; border-radius: 16px; padding: 18px 20px;
  border: 1px solid var(--am-line); box-shadow: 0 6px 18px rgba(15,23,42,0.06);
  margin-bottom: 16px;
}
.am-matrix-title { font-weight: 750; color: var(--am-ink); margin-bottom: 12px; }
.am-matrix-grid { display: grid; grid-template-columns: 1fr 1fr; gap: 10px; }
.am-quad { border-radius: 12px; padding: 14px 16px; min-height: 78px; border: 1px solid transparent; }
.am-quad .am-quad-h { font-weight: 750; font-size: 0.98rem; }
.am-quad .am-quad-s { font-size: 0.82rem; opacity: 0.85; margin-top: 2px; }
.am-quad.green { background: #ecfdf5; border-color: #6ee7b7; color: #065f46; }
.am-quad.amber { background: #fffbeb; border-color: #fcd34d; color: #92400e; }
.am-quad.red { background: #fef2f2; border-color: #fca5a5; color: #991b1b; }
.am-quad.slate { background: #f8fafc; border-color: #e2e8f0; color: #475569; }
.am-matrix-axes {
  display: flex; justify-content: space-between; margin-top: 10px;
  font-size: 0.76rem; color: var(--am-muted); font-weight: 700;
  text-transform: uppercase; letter-spacing: 0.06em;
}
</style>
"""
st.markdown(CSS, unsafe_allow_html=True)


# ---------------------------------------------------------------------------
# Session state defaults
# ---------------------------------------------------------------------------
def init_state():
    ss = st.session_state
    ss.setdefault("nav", "Session Progress")
    ss.setdefault("orders", {})
    ss.setdefault("missions_completed", {})       # idx -> True
    ss.setdefault("active_mission", None)
    ss.setdefault("mission_answered", {})         # idx -> bool
    ss.setdefault("mission_selected", {})         # idx -> str
    ss.setdefault("mission_correct", {})          # idx -> bool
    # Practice deck
    ss.setdefault("practice_deck", None)          # list of row indices
    ss.setdefault("practice_pos", 0)
    ss.setdefault("practice_sig", None)
    ss.setdefault("practice_answered", {})        # row idx -> bool
    ss.setdefault("practice_selected", {})        # row idx -> str
    ss.setdefault("practice_correct", {})         # row idx -> bool
    ss.setdefault("practice_history", [])         # list of dicts
    ss.setdefault("audit_passed", None)


init_state()

DATA = load_data(DATA_PATH)
SCEN = DATA["Scenarios"]
PRINCIPLES = DATA["Access Model Principles"]
PRINCIPLE_LOOKUP = {str(r["Principle"]).strip(): r for _, r in PRINCIPLES.iterrows()}

# Run the audit guardrail once and record the result.
if st.session_state["audit_passed"] is None:
    _violations = audit_scenarios(SCEN)
    st.session_state["audit_passed"] = (len(_violations) == 0)
    st.session_state["audit_violation_count"] = len(_violations)


# ---------------------------------------------------------------------------
# Small helpers
# ---------------------------------------------------------------------------
def goto(page):
    st.session_state["nav"] = page
    st.rerun()


def principle_def(name):
    row = PRINCIPLE_LOOKUP.get(str(name).strip())
    return str(row["Plain_Language_Definition"]) if row is not None else ""


def principle_trap(name):
    row = PRINCIPLE_LOOKUP.get(str(name).strip())
    return str(row["Common_Misinterpretation"]) if row is not None else ""


# ---------------------------------------------------------------------------
# Inline SVG graphics (no external assets). All visuals render via CSS/SVG.
# ---------------------------------------------------------------------------
ICON_PATHS = {
    "compass": '<circle cx="12" cy="12" r="9"/><polygon points="15.5,8.5 11,11 8.5,15.5 13,13" '
               'fill="currentColor" stroke="none"/>',
    "shield": '<path d="M12 3l7 3v5c0 4.5-3 7.3-7 8.5C8 18.3 5 15.5 5 11V6z"/>',
    "lifebuoy": '<circle cx="12" cy="12" r="9"/><circle cx="12" cy="12" r="3.4"/>'
                '<line x1="12" y1="3" x2="12" y2="8.6"/><line x1="12" y1="15.4" x2="12" y2="21"/>'
                '<line x1="3" y1="12" x2="8.6" y2="12"/><line x1="15.4" y1="12" x2="21" y2="12"/>',
    "users": '<circle cx="9" cy="9" r="3"/><path d="M3.5 19a5.5 5.5 0 0 1 11 0"/>'
             '<circle cx="17" cy="9.5" r="2.3"/><path d="M15 14.4a5 5 0 0 1 5.5 4.6"/>',
    "loop": '<path d="M4 12a8 8 0 0 1 13.7-5.6L20 8"/><polyline points="20 3 20 8 15 8"/>'
            '<path d="M20 12a8 8 0 0 1-13.7 5.6L4 16"/><polyline points="4 21 4 16 9 16"/>',
    "network": '<circle cx="12" cy="5" r="2.2"/><circle cx="5" cy="18" r="2.2"/>'
               '<circle cx="19" cy="18" r="2.2"/><line x1="12" y1="7" x2="6" y2="16"/>'
               '<line x1="12" y1="7" x2="18" y2="16"/><line x1="7.2" y1="18" x2="16.8" y2="18"/>',
    "bars": '<rect x="4" y="13" width="3.6" height="7" rx="1" fill="currentColor" stroke="none"/>'
            '<rect x="10.2" y="9" width="3.6" height="11" rx="1" fill="currentColor" stroke="none"/>'
            '<rect x="16.4" y="5" width="3.6" height="15" rx="1" fill="currentColor" stroke="none"/>',
    "clipboard": '<rect x="6" y="4" width="12" height="17" rx="2"/>'
                 '<rect x="9" y="2.5" width="6" height="3" rx="1"/><line x1="9" y1="10" x2="15" y2="10"/>'
                 '<line x1="9" y1="13.5" x2="15" y2="13.5"/><line x1="9" y1="17" x2="13" y2="17"/>',
    "rocket": '<path d="M12 3c3 1.2 5 4.2 5 8l-2.4 2.4H9.4L7 11c0-3.8 2-6.8 5-8z"/>'
              '<circle cx="12" cy="9.6" r="1.5"/><path d="M9.4 16c-1 1.4-1 3.4-1 4 .6 0 2.6 0 4-1"/>',
    "flag": '<line x1="5" y1="3" x2="5" y2="21"/><path d="M5 4h12l-2.5 4L17 12H5z"/>',
    "target": '<circle cx="12" cy="12" r="9"/><circle cx="12" cy="12" r="5"/>'
              '<circle cx="12" cy="12" r="1.5" fill="currentColor" stroke="none"/>',
    "bulb": '<path d="M9 18h6"/><path d="M10 21h4"/>'
            '<path d="M12 3a6 6 0 0 0-4 10.5c.8.7 1 1.4 1 2.5h6c0-1.1.2-1.8 1-2.5A6 6 0 0 0 12 3z"/>',
    "swap": '<polyline points="7 4 3 8 7 12"/><line x1="3" y1="8" x2="16" y2="8"/>'
            '<polyline points="17 12 21 16 17 20"/><line x1="21" y1="16" x2="8" y2="16"/>',
    "trophy": '<path d="M8 4h8v4a4 4 0 0 1-8 0z"/><path d="M8 5.5H5.5a2.5 2.5 0 0 0 2.5 2.5"/>'
              '<path d="M16 5.5h2.5a2.5 2.5 0 0 1-2.5 2.5"/><line x1="12" y1="12" x2="12" y2="15"/>'
              '<path d="M9 20h6l-1-4h-4z"/>',
    "spark": '<path d="M12 3l1.8 5.2L19 10l-5.2 1.8L12 17l-1.8-5.2L5 10l5.2-1.8z" '
             'fill="currentColor" stroke="none"/>',
}

CATEGORY_ICONS = {
    "Self-Regulation": "lifebuoy",
    "Repair Protocols": "loop",
    "Timeback": "bars",
    "Check Charts": "clipboard",
    "Launch Facilitation": "rocket",
    "Experiences": "flag",
    "Teacher Coaching": "users",
    "Principal Concerns": "shield",
    "Public School Compliance": "clipboard",
    "Multilingual Learners": "network",
    "Special Education": "lifebuoy",
    "Attendance": "flag",
}

PRINCIPLE_ICONS = {
    "Autonomy Is the Currency": "compass",
    "Capability Through Friction": "bars",
    "Safety First": "shield",
    "Support, Don't Solve": "lifebuoy",
    "Belonging Through Contribution": "users",
    "If It Doesn't Work, It Isn't Finished": "loop",
    "Knowledge Moves in All Directions": "network",
}


def icon_svg(name, color=DEFAULT_ACCENT, size=22):
    paths = ICON_PATHS.get(name, ICON_PATHS["bulb"])
    return (f'<svg viewBox="0 0 24 24" width="{size}" height="{size}" fill="none" '
            f'stroke="{color}" stroke-width="1.8" stroke-linecap="round" '
            f'stroke-linejoin="round" style="color:{color}">{paths}</svg>')


def category_icon(category):
    return CATEGORY_ICONS.get(str(category), "bulb")


def principle_icon(principle):
    return PRINCIPLE_ICONS.get(str(principle).strip(), "spark")


def donut(pct, color=DEFAULT_ACCENT, size=140, center_text=None, sub=None):
    """Return an SVG progress ring as an HTML string."""
    pct = max(0.0, min(100.0, float(pct)))
    r, cx = 52, 65
    circ = 2 * math.pi * r
    dash = circ * pct / 100.0
    text = center_text if center_text is not None else f"{round(pct)}%"
    sub_html = f'<div class="am-donut-sub">{sub}</div>' if sub else ""
    return (
        f'<div class="am-donut-wrap"><svg width="{size}" height="{size}" viewBox="0 0 130 130">'
        f'<circle cx="{cx}" cy="{cx}" r="{r}" fill="none" stroke="#e2e8f0" stroke-width="12"/>'
        f'<circle cx="{cx}" cy="{cx}" r="{r}" fill="none" stroke="{color}" stroke-width="12" '
        f'stroke-linecap="round" stroke-dasharray="{dash:.2f} {circ:.2f}" '
        f'transform="rotate(-90 {cx} {cx})"/>'
        f'<text x="{cx}" y="{cx}" text-anchor="middle" dominant-baseline="central" '
        f'font-size="26" font-weight="800" fill="#0f172a">{text}</text></svg>{sub_html}</div>'
    )


def support_expectations_matrix():
    """A 2x2 graphic showing the Access Model lives at high support + high expectations."""
    return (
        '<div class="am-matrix-card">'
        '<div class="am-matrix-title">Where the Access Model lives</div>'
        '<div class="am-matrix-grid">'
        '<div class="am-quad amber"><div class="am-quad-h">Rescuing</div>'
        '<div class="am-quad-s">High support, low expectations</div></div>'
        '<div class="am-quad green"><div class="am-quad-h">Access Model</div>'
        '<div class="am-quad-s">High support, high expectations</div></div>'
        '<div class="am-quad slate"><div class="am-quad-h">Disengaged</div>'
        '<div class="am-quad-s">Low support, low expectations</div></div>'
        '<div class="am-quad red"><div class="am-quad-h">Pressure</div>'
        '<div class="am-quad-s">Low support, high expectations</div></div>'
        '</div>'
        '<div class="am-matrix-axes"><span>&#8593; More support</span>'
        '<span>More expectations &#8594;</span></div>'
        '</div>'
    )


def card(title, summary, accent=DEFAULT_ACCENT, body_html="", icon_name=None,
         icon_color=None):
    if icon_name:
        ic = icon_svg(icon_name, icon_color or accent, size=24)
        head = (f'<div class="am-card-head"><span class="am-icon">{ic}</span>'
                f'<span class="am-card-title">{title}</span></div>')
    else:
        head = f'<div class="am-card-title">{title}</div>'
    html = (
        f'<div class="am-card" style="border-left-color:{accent}">'
        f'{head}'
        f'<div class="am-card-summary">{summary}</div>'
        f'{body_html}'
        f'</div>'
    )
    st.markdown(html, unsafe_allow_html=True)


def stat_card(label, value, sub=""):
    st.markdown(
        f'<div class="am-stat"><div class="am-stat-label">{label}</div>'
        f'<div class="am-stat-value">{value}</div>'
        f'<div class="am-stat-sub">{sub}</div></div>',
        unsafe_allow_html=True,
    )


def get_scenario_row(category, topic):
    m = SCEN[(SCEN["Category"] == category) & (SCEN["Topic"] == topic)]
    if len(m) == 0:
        return None
    return m.iloc[0]


def completed_count():
    return sum(1 for i in range(TOTAL_MISSIONS)
               if st.session_state["missions_completed"].get(i))


def mission_unlocked(idx):
    if idx == 0:
        return True
    return bool(st.session_state["missions_completed"].get(idx - 1))


def level_missions(level):
    return [i for i, m in enumerate(MISSIONS) if m["level"] == level]


def level_complete(level):
    return all(st.session_state["missions_completed"].get(i)
               for i in level_missions(level))


def level_unlocked(level):
    if level == 1:
        return True
    return level_complete(level - 1)


def certification_level_label():
    done = completed_count()
    if done >= TOTAL_MISSIONS:
        return "Certified"
    if done >= 6:
        return "Level 3"
    if done >= 3:
        return "Level 2"
    if done >= 1:
        return "Level 1"
    return "Not started"


# ===========================================================================
# Sidebar navigation
# ===========================================================================
# Resource material (REFERENCE) is presented before the mission modules (LEARN)
# so Guides build foundation before practicing decisions.
NAV_GROUPS = [
    ("START HERE", [("Foundations", "Foundations"), ("Guide Role", "Guide Role")]),
    ("LEARN", [("Missions", "Missions"), ("Guide Certification", "Guide Certification")]),
    ("PRACTICE", [("Scenario Challenge", "Scenario Challenge")]),
    ("IMPLEMENT", [
        ("Launch Toolkit", "Launch Toolkit"),
        ("Check Charts", "Check Charts"),
        ("Repair Protocols", "Repair Protocols"),
        ("Alpha → Access", "Alpha → Access"),
    ]),
    ("PROGRESS", [
        ("Session Progress", "Session Progress"),
        ("Missions Completed", "Missions Completed"),
        ("Scenarios Completed", "Scenarios Completed"),
    ]),
]


def render_sidebar():
    with st.sidebar:
        st.markdown(
            '<div class="am-brand"><div class="am-brand-main">ACCESS MODEL</div>'
            '<div class="am-brand-sub">Operating System</div></div>',
            unsafe_allow_html=True,
        )
        active = st.session_state["nav"]
        if st.button("Home", key="nav_home",
                     type="primary" if active == "Session Progress" else "secondary",
                     use_container_width=True):
            goto("Session Progress")
        for group, items in NAV_GROUPS:
            st.markdown(f'<div class="am-navgroup">{group}</div>', unsafe_allow_html=True)
            for label, page in items:
                btype = "primary" if active == page else "secondary"
                if st.button(label, key=f"nav_{page}", type=btype, use_container_width=True):
                    goto(page)


# ===========================================================================
# PROGRESS / dashboard pages
# ===========================================================================
def page_dashboard():
    st.markdown(
        '<div class="am-hero"><h1>Welcome back!</h1>'
        '<p>Your Future 2 Guide training command center. Learn the Access Model '
        'through decisions, practice on real scenarios, and put it to work.</p></div>',
        unsafe_allow_html=True,
    )

    done = completed_count()
    answered = len(st.session_state["practice_history"])
    correct = sum(1 for h in st.session_state["practice_history"] if h["correct"])
    accuracy = f"{round(100 * correct / answered)}%" if answered else "--"

    c1, c2, c3, c4 = st.columns(4)
    with c1:
        stat_card("Missions completed", f"{done}/{TOTAL_MISSIONS}", "Guide certification")
    with c2:
        stat_card("Certification level", certification_level_label(),
                  "Across 3 levels")
    with c3:
        stat_card("Scenarios answered", str(answered), "Practice challenge")
    with c4:
        stat_card("Accuracy", accuracy, f"{correct} aligned" if answered else "No answers yet")

    st.write("")
    st.markdown('<div class="am-section">Recommended path</div>', unsafe_allow_html=True)
    st.markdown('<div class="am-sub">Start with the resource material, then move into '
                'the mission modules.</div>', unsafe_allow_html=True)
    s1, s2, s3 = st.columns(3)
    with s1:
        card("Step 1 &middot; Foundations",
             "Understand the thinking behind the Access Model.", accent="#3b82f6")
        if st.button("Open Foundations", key="path_foundations", use_container_width=True):
            goto("Foundations")
    with s2:
        card("Step 2 &middot; Guide Role",
             "See the seven principles through the Guide's lens.", accent="#6366f1")
        if st.button("Open Guide Role", key="path_guiderole", use_container_width=True):
            goto("Guide Role")
    with s3:
        card("Step 3 &middot; Missions",
             "Apply it in decision-based mission modules.", accent="#8b5cf6")
        if st.button("Begin missions", type="primary", key="path_missions",
                     use_container_width=True):
            goto("Missions")

    st.write("")
    left, right = st.columns([2, 1])
    with left:
        pct = round(100 * done / TOTAL_MISSIONS)
        card("Mission progress",
             f"You have completed {done} of {TOTAL_MISSIONS} missions.",
             accent="#6366f1", icon_name="trophy")
        d1, d2 = st.columns([1, 1.3])
        with d1:
            st.markdown(donut(pct, color="#6366f1", sub="Certification"),
                        unsafe_allow_html=True)
        with d2:
            st.write("")
            st.progress(pct / 100)
            st.caption(f"Level: {certification_level_label()}")
            if st.button("Continue missions", type="primary", key="dash_continue"):
                goto("Missions")
    with right:
        card("Quick actions", "Jump straight into training.", accent="#ec4899",
             icon_name="spark")
        if st.button("Scenario Challenge", key="dash_practice", use_container_width=True):
            goto("Scenario Challenge")
        if st.button("Guide Certification", key="dash_cert", use_container_width=True):
            goto("Guide Certification")

    st.write("")
    st.markdown('<div class="am-section">Session activity</div>', unsafe_allow_html=True)
    a1, a2, a3 = st.columns(3)
    with a1:
        stat_card("Current level", certification_level_label(), "")
    with a2:
        next_idx = next((i for i in range(TOTAL_MISSIONS)
                         if not st.session_state["missions_completed"].get(i)), None)
        next_label = MISSIONS[next_idx]["title"] if next_idx is not None else "All complete"
        stat_card("Next mission", next_label, "")
    with a3:
        revise = sum(1 for h in st.session_state["practice_history"] if not h["correct"])
        stat_card("Needs revision", str(revise), "Practice answers to revisit")

    st.write("")
    st.markdown('<div class="am-section">Scenario bank at a glance</div>',
                unsafe_allow_html=True)
    st.caption("The breadth of situations this training covers.")
    g1, g2 = st.columns(2)
    with g1:
        st.markdown("**By category**")
        cat_counts = (SCEN["Category"].astype(str).value_counts()
                      .rename_axis("Category").to_frame("Scenarios"))
        st.bar_chart(cat_counts, horizontal=True, color="#8b5cf6")
    with g2:
        st.markdown("**By grade band**")
        band_counts = (SCEN["Age_Band"].map(grade_label)
                       .value_counts().rename_axis("Grade band").to_frame("Scenarios"))
        st.bar_chart(band_counts, color="#0ea5e9")

    if st.session_state["audit_passed"] is False:
        st.error("Answer-choice audit found banned content. Review distractor rules.")


def page_missions_completed():
    st.markdown('<div class="am-section">Missions Completed</div>', unsafe_allow_html=True)
    st.markdown(f'<div class="am-sub">{completed_count()} of {TOTAL_MISSIONS} missions complete.</div>',
                unsafe_allow_html=True)
    for level in LEVELS:
        st.markdown(f"**Level {level}**")
        for i in level_missions(level):
            m = MISSIONS[i]
            if st.session_state["missions_completed"].get(i):
                badge = '<span class="am-badge done">Complete</span>'
            elif mission_unlocked(i):
                badge = '<span class="am-badge progress">Available</span>'
            else:
                badge = '<span class="am-badge locked">Locked</span>'
            card(f"{m['title']} {badge}",
                 f"{m['category']} &middot; {m['topic']}", accent=accent_for(m["category"]),
                 icon_name=category_icon(m["category"]))
    if st.button("Reset missions", key="reset_missions"):
        st.session_state["missions_completed"] = {}
        st.session_state["mission_answered"] = {}
        st.session_state["mission_selected"] = {}
        st.session_state["mission_correct"] = {}
        st.session_state["active_mission"] = None
        st.rerun()


def page_scenarios_completed():
    st.markdown('<div class="am-section">Scenarios Completed</div>', unsafe_allow_html=True)
    history = st.session_state["practice_history"]
    if not history:
        st.info("No practice scenarios answered yet. Try the Scenario Challenge.")
    else:
        correct = sum(1 for h in history if h["correct"])
        st.markdown(
            f'<div class="am-sub">{len(history)} answered &middot; {correct} aligned &middot; '
            f'{len(history) - correct} need revision.</div>', unsafe_allow_html=True)
        for h in reversed(history):
            badge = ('<span class="am-badge good">Aligned</span>' if h["correct"]
                     else '<span class="am-badge revise">Needs revision</span>')
            card(f"{h['topic']} {badge}",
                 f"{h['category']} &middot; {grade_label(h['age_band'])}",
                 accent=accent_for(h["category"]), icon_name=category_icon(h["category"]),
                 body_html=f'<div class="am-card-body">{h["scenario"]}</div>')
    if st.button("Reset practice history", key="reset_hist"):
        st.session_state["practice_history"] = []
        st.rerun()


# ===========================================================================
# LEARN: Missions
# ===========================================================================
def coaching_block(row):
    principle = str(row["Access_Model_Principle"]).strip()
    st.markdown('<div class="am-section" style="font-size:1.05rem;">Coaching</div>',
                unsafe_allow_html=True)
    card("Related principle", f"{principle}",
         accent=accent_for(row["Category"]), icon_name=principle_icon(principle),
         body_html=f'<div class="am-card-body">{principle_def(principle)}</div>')
    card("Strong Guide move", "What an aligned Guide does next.",
         accent="#10b981",
         body_html=f'<div class="am-card-body">{row["Strong_Response"]}</div>')
    card("Common trap", "A well-intentioned mistake to avoid.",
         accent="#ef4444",
         body_html=f'<div class="am-card-body">{principle_trap(principle)}</div>')


def render_graded_options(options, correct, selected):
    for opt in options:
        if opt == correct:
            cls = "correct"
            tag = " (Strong Guide move)"
        elif opt == selected:
            cls = "wrong"
            tag = " (Your choice)"
        else:
            cls = "muted"
            tag = ""
        st.markdown(f'<div class="am-opt {cls}">{opt}{tag}</div>', unsafe_allow_html=True)


def run_mission(idx):
    m = MISSIONS[idx]
    row = get_scenario_row(m["category"], m["topic"])
    if st.button("← Back to missions", key="mission_back"):
        st.session_state["active_mission"] = None
        st.rerun()

    if row is None:
        st.error(f"Scenario not found for {m['category']} / {m['topic']}.")
        return

    accent = accent_for(m["category"])
    is_done = bool(st.session_state["missions_completed"].get(idx))
    st.markdown(
        f'<div class="am-section">Level {m["level"]} &middot; {m["title"]}</div>',
        unsafe_allow_html=True)
    st.markdown(
        f'<span class="am-chip">{m["category"]}</span>'
        f'<span class="am-chip">{grade_label(row["Age_Band"])}</span>',
        unsafe_allow_html=True)
    st.write("")
    card("Scenario", row["Scenario"], accent=accent,
         icon_name=category_icon(row["Category"]))

    correct, over, under = make_choices(row)
    options = ordered_choices(f"mission:{idx}", [correct, over, under])
    answered = st.session_state["mission_answered"].get(idx, False)

    # Completed missions are revisited read-only.
    if is_done and not answered:
        st.success("Mission complete. Reviewing in read-only mode.")
        render_graded_options(options, correct, correct)
        coaching_block(row)
        nxt = idx + 1
        if nxt < TOTAL_MISSIONS and mission_unlocked(nxt):
            if st.button("Go to next mission", type="primary", key="mission_next_ro"):
                st.session_state["active_mission"] = nxt
                st.rerun()
        return

    if not answered:
        st.markdown("**As the Guide, what should happen next?**")
        for i, opt in enumerate(options):
            if st.button(opt, key=f"m_{idx}_opt_{i}", use_container_width=True):
                st.session_state["mission_answered"][idx] = True
                st.session_state["mission_selected"][idx] = opt
                st.session_state["mission_correct"][idx] = (opt == correct)
                st.rerun()
        return

    selected = st.session_state["mission_selected"].get(idx)
    is_correct = st.session_state["mission_correct"].get(idx, False)
    if is_correct:
        st.success("Correct. That is the aligned Guide move.")
    else:
        st.error("Not yet. That move does not fit the Access Model here.")
    render_graded_options(options, correct, selected)
    coaching_block(row)

    if is_correct:
        if not is_done:
            if st.button("Complete mission and continue", type="primary",
                         key=f"m_{idx}_complete"):
                st.session_state["missions_completed"][idx] = True
                nxt = idx + 1
                st.session_state["active_mission"] = nxt if nxt < TOTAL_MISSIONS else None
                st.rerun()
        else:
            nxt = idx + 1
            if nxt < TOTAL_MISSIONS and mission_unlocked(nxt):
                if st.button("Go to next mission", type="primary", key=f"m_{idx}_next"):
                    st.session_state["active_mission"] = nxt
                    st.rerun()
    else:
        if st.button("Try again", key=f"m_{idx}_retry"):
            st.session_state["mission_answered"][idx] = False
            st.rerun()


def page_missions():
    if st.session_state["active_mission"] is not None:
        run_mission(st.session_state["active_mission"])
        return

    st.markdown('<div class="am-section">Missions</div>', unsafe_allow_html=True)
    st.markdown(
        '<div class="am-sub">Nine missions across three certification levels. '
        'One scenario, one decision at a time. Each mission unlocks when you complete '
        'the one before it.</div>', unsafe_allow_html=True)

    done = completed_count()
    pct = round(100 * done / TOTAL_MISSIONS)
    p1, p2 = st.columns([1, 2.4])
    with p1:
        st.markdown(donut(pct, color="#8b5cf6", sub=f"{done}/{TOTAL_MISSIONS} missions"),
                    unsafe_allow_html=True)
    with p2:
        st.write("")
        st.progress(pct / 100)
        st.caption(f"{done} of {TOTAL_MISSIONS} missions complete &middot; "
                   f"Level: {certification_level_label()}")

    for level in LEVELS:
        locked_level = not level_unlocked(level)
        suffix = " (locked)" if locked_level else ""
        st.markdown(f"### Level {level}{suffix}")
        for i in level_missions(level):
            m = MISSIONS[i]
            unlocked = mission_unlocked(i)
            is_done = bool(st.session_state["missions_completed"].get(i))
            if is_done:
                badge = '<span class="am-badge done">Complete</span>'
            elif unlocked:
                badge = '<span class="am-badge progress">In progress</span>'
            else:
                badge = '<span class="am-badge locked">Locked</span>'
            row = get_scenario_row(m["category"], m["topic"])
            grade = grade_label(row["Age_Band"]) if row is not None else ""
            card(f"Mission {i + 1}: {m['title']} {badge}",
                 f"{m['category']} &middot; {m['topic']} &middot; {grade}",
                 accent=accent_for(m["category"]),
                 icon_name=category_icon(m["category"]))
            if is_done:
                if st.button("Review mission", key=f"open_{i}"):
                    st.session_state["active_mission"] = i
                    st.rerun()
            elif unlocked:
                if st.button("Start mission", type="primary", key=f"open_{i}"):
                    st.session_state["active_mission"] = i
                    st.rerun()
            else:
                st.button("Locked", key=f"open_{i}", disabled=True)


def page_certification():
    st.markdown('<div class="am-section">Future 2 Guide Certification</div>',
                unsafe_allow_html=True)
    done = completed_count()
    pct = round(100 * done / TOTAL_MISSIONS)
    st.markdown(f'<div class="am-sub">Overall progress: {done}/{TOTAL_MISSIONS} '
                f'missions ({pct}%).</div>', unsafe_allow_html=True)
    o1, o2 = st.columns([1, 2.4])
    with o1:
        ring_color = "#10b981" if done >= TOTAL_MISSIONS else "#8b5cf6"
        st.markdown(donut(pct, color=ring_color, sub="Certified" if done >= TOTAL_MISSIONS
                          else certification_level_label()), unsafe_allow_html=True)
    with o2:
        st.write("")
        st.progress(pct / 100)
        st.caption(f"{done} of {TOTAL_MISSIONS} missions complete across three levels.")

    if done >= TOTAL_MISSIONS:
        st.success("Certified. You have completed all nine missions across three levels.")

    for level in LEVELS:
        miss = level_missions(level)
        ldone = sum(1 for i in miss if st.session_state["missions_completed"].get(i))
        if level_complete(level):
            badge = '<span class="am-badge done">Complete</span>'
        elif level_unlocked(level):
            badge = '<span class="am-badge progress">In progress</span>'
        else:
            badge = '<span class="am-badge locked">Locked</span>'
        checklist = ""
        for i in miss:
            mark = "[x]" if st.session_state["missions_completed"].get(i) else "[ ]"
            checklist += f'<div class="am-kv">{mark} {MISSIONS[i]["title"]}</div>'
        card(f"Level {level} {badge}",
             f"{ldone} of {len(miss)} missions complete.",
             accent="#8b5cf6", body_html=checklist,
             icon_name="trophy" if level_complete(level) else "target")
        st.progress(ldone / len(miss))

    if st.button("Continue training", type="primary", key="cert_continue"):
        goto("Missions")


# ===========================================================================
# PRACTICE: Scenario Challenge
# ===========================================================================
def practice_filter_options():
    bands = ["All grades"] + [grade_label(b) for b in
                              ["3-4", "5-6", "7-8", "All", "Adult"]
                              if b in set(SCEN["Age_Band"].astype(str))]
    cats = ["All categories"] + sorted(SCEN["Category"].astype(str).unique().tolist())
    return bands, cats


LABEL_TO_BAND = {grade_label(b): b for b in ["3-4", "5-6", "7-8", "All", "Adult"]}


def build_deck(band_label, category):
    df = SCEN
    if band_label != "All grades":
        df = df[df["Age_Band"].astype(str) == LABEL_TO_BAND[band_label]]
    if category != "All categories":
        df = df[df["Category"].astype(str) == category]
    return df.index.tolist()


def page_practice():
    st.markdown('<div class="am-section">Scenario Challenge</div>', unsafe_allow_html=True)
    st.markdown('<div class="am-sub">Work through a shuffled deck of real scenarios. '
                'Choose the aligned Guide move and get instant coaching.</div>',
                unsafe_allow_html=True)

    bands, cats = practice_filter_options()
    f1, f2 = st.columns(2)
    with f1:
        band_label = st.selectbox("Grade band", bands, key="practice_band")
    with f2:
        category = st.selectbox("Category", cats, key="practice_cat")

    sig = (band_label, category)
    if st.session_state["practice_sig"] != sig or st.session_state["practice_deck"] is None:
        st.session_state["practice_sig"] = sig
        st.session_state["practice_deck"] = build_deck(band_label, category)
        st.session_state["practice_pos"] = 0

    deck = st.session_state["practice_deck"]
    n = len(deck)
    if n == 0:
        st.warning("No scenarios match these filters. Try a different combination.")
        return

    # Live metrics from this deck.
    answered = sum(1 for ri in deck if st.session_state["practice_answered"].get(ri))
    correct = sum(1 for ri in deck if st.session_state["practice_correct"].get(ri))
    acc_val = round(100 * correct / answered) if answered else 0
    acc = f"{acc_val}%" if answered else "--"
    mcols = st.columns([1.1, 1, 1, 1])
    with mcols[0]:
        ring = "#10b981" if acc_val >= 70 else ("#f59e0b" if answered else "#94a3b8")
        st.markdown(donut(acc_val, color=ring, center_text=acc, sub="Accuracy", size=120),
                    unsafe_allow_html=True)
    with mcols[1]:
        stat_card("Correct", str(correct), "Aligned moves")
    with mcols[2]:
        stat_card("Answered", f"{answered}/{n}", "This deck")
    with mcols[3]:
        stat_card("Needs revision", str(answered - correct), "Revisit these")

    st.write("")
    pos = st.session_state["practice_pos"]
    if pos >= n:
        st.success(f"Deck complete. You answered {answered} scenarios "
                   f"with {acc} accuracy.")
        if st.button("Reset practice", key="practice_reset_done"):
            reset_practice()
            st.rerun()
        return

    ri = deck[pos]
    row = SCEN.loc[ri]
    accent = accent_for(row["Category"])
    st.caption(f"Scenario {pos + 1} of {n}")
    st.markdown(
        f'<span class="am-chip">{row["Category"]}</span>'
        f'<span class="am-chip">{grade_label(row["Age_Band"])}</span>'
        f'<span class="am-chip">{row["Topic"]}</span>', unsafe_allow_html=True)
    st.write("")
    card("Scenario", row["Scenario"], accent=accent,
         icon_name=category_icon(row["Category"]))

    correct_ans, over, under = make_choices(row)
    options = ordered_choices(f"practice:{ri}", [correct_ans, over, under])
    is_answered = st.session_state["practice_answered"].get(ri, False)

    if not is_answered:
        st.markdown("**As the Guide, what should happen next?**")
        for i, opt in enumerate(options):
            if st.button(opt, key=f"p_{ri}_opt_{i}", use_container_width=True):
                is_corr = (opt == correct_ans)
                st.session_state["practice_answered"][ri] = True
                st.session_state["practice_selected"][ri] = opt
                st.session_state["practice_correct"][ri] = is_corr
                st.session_state["practice_history"].append({
                    "category": str(row["Category"]),
                    "topic": str(row["Topic"]),
                    "age_band": str(row["Age_Band"]),
                    "scenario": str(row["Scenario"]),
                    "correct": is_corr,
                })
                st.rerun()
    else:
        selected = st.session_state["practice_selected"].get(ri)
        is_corr = st.session_state["practice_correct"].get(ri, False)
        if is_corr:
            st.success("Correct. This move is aligned.")
        else:
            st.error("Needs revision. This move is not aligned here.")
        render_graded_options(options, correct_ans, selected)
        principle = str(row["Access_Model_Principle"]).strip()
        card("Why this is" + ("" if is_corr else " not") + " aligned",
             "Feedback on the Access Model fit.", accent=accent,
             body_html=f'<div class="am-card-body">{row["Feedback"]}</div>')
        card("Stronger Guide move", "The aligned response.", accent="#10b981",
             body_html=f'<div class="am-card-body">{row["Strong_Response"]}</div>')
        card("Related principle", principle, accent="#6366f1",
             body_html=f'<div class="am-card-body">{principle_def(principle)}</div>')
        c1, c2 = st.columns(2)
        with c1:
            if st.button("Next scenario", type="primary", key=f"p_{ri}_next"):
                st.session_state["practice_pos"] = pos + 1
                st.rerun()
        with c2:
            if st.button("Reset practice", key=f"p_{ri}_reset"):
                reset_practice()
                st.rerun()


def reset_practice():
    st.session_state["practice_deck"] = build_deck(*st.session_state["practice_sig"]) \
        if st.session_state["practice_sig"] else None
    st.session_state["practice_pos"] = 0
    st.session_state["practice_answered"] = {}
    st.session_state["practice_selected"] = {}
    st.session_state["practice_correct"] = {}
    st.session_state["practice_history"] = []


# ===========================================================================
# IMPLEMENT toolkit
# ===========================================================================
def page_repair():
    st.markdown('<div class="am-section">Repair Protocols</div>', unsafe_allow_html=True)
    st.markdown('<div class="am-sub">Harm is addressed through repair and re-entry, '
                'restoring belonging through contribution.</div>', unsafe_allow_html=True)
    guidance = [
        ("Individual harm", "Individual harm requires individual repair before re-entry."),
        ("Group harm", "Group harm requires group repair before re-entry."),
        ("Property harm", "Property harm requires restitution and repair before re-entry."),
    ]
    cols = st.columns(3)
    for col, (title, text) in zip(cols, guidance):
        with col:
            card(title, text, accent="#ec4899", icon_name="loop")

    st.markdown('<div class="am-section" style="font-size:1.1rem;">Related scenarios</div>',
                unsafe_allow_html=True)
    rdf = SCEN[SCEN["Category"] == "Repair Protocols"]
    for _, row in rdf.iterrows():
        principle = str(row["Access_Model_Principle"]).strip()
        body = (
            f'<div class="am-kv"><b>Recommended response:</b> {row["Strong_Response"]}</div>'
            f'<div class="am-kv"><b>Why it is aligned:</b> {row["Feedback"]}</div>'
            f'<div class="am-kv"><b>Related principle:</b> {principle}</div>'
        )
        card(f'{row["Topic"]}', row["Scenario"], accent=accent_for("Repair Protocols"),
             body_html=body, icon_name="loop")


def page_launch_toolkit():
    st.markdown('<div class="am-section">Launch Toolkit</div>', unsafe_allow_html=True)
    st.markdown('<div class="am-sub">Daily Launch lessons that build regulation and '
                'culture skills.</div>', unsafe_allow_html=True)
    df = DATA["Launch Lessons"]
    bands = ["All grades"] + [grade_label(b) for b in ["3-4", "5-6", "7-8", "All", "Adult"]
                              if b in set(df["Age_Band"].astype(str))]
    pick = st.selectbox("Grade band", bands, key="launch_band")
    view = df
    if pick != "All grades":
        view = df[df["Age_Band"].astype(str) == LABEL_TO_BAND[pick]]
    if len(view) == 0:
        st.info("No lessons for this grade band.")
    for _, row in view.iterrows():
        body = (
            f'<div class="am-kv"><b>Skill:</b> {row["Self_Regulation_or_Culture_Skill"]}</div>'
            f'<div class="am-kv"><b>Grade band:</b> {grade_label(row["Age_Band"])} '
            f'&middot; <b>Duration:</b> {row["Duration"]}</div>'
        )
        card(row["Lesson_Title"], row["Lesson_Description"], accent="#f59e0b",
             body_html=body, icon_name="rocket")
        with st.expander("Expand to Learn More"):
            st.markdown(f"**Mastery evidence:** {row['Mastery_Evidence']}")
            st.markdown(f"**Aligned categories:** {row['Aligned_Categories']}")


def page_check_charts():
    st.markdown('<div class="am-section">Check Charts</div>', unsafe_allow_html=True)
    st.markdown('<div class="am-sub">Student-owned checks where evidence, not '
                'completion, defines mastery.</div>', unsafe_allow_html=True)
    df = DATA["Check Chart Alignments"]
    areas = ["All skill areas"] + sorted(df["Skill_Area"].astype(str).unique().tolist())
    pick = st.selectbox("Skill area", areas, key="check_area")
    view = df if pick == "All skill areas" else df[df["Skill_Area"].astype(str) == pick]
    for _, row in view.iterrows():
        body = (
            f'<div class="am-kv"><b>Skill area:</b> {row["Skill_Area"]}</div>'
            f'<div class="am-kv"><b>Evidence required:</b> {row["Evidence_Required"]}</div>'
        )
        card(row["Check_Name"], row["Check_Description"], accent="#14b8a6",
             body_html=body, icon_name="clipboard")
        with st.expander("Expand to Learn More"):
            st.markdown(f"**Mastery criteria:** {row['Mastery_Criteria']}")
            st.markdown(f"**Aligned categories:** {row['Aligned_Categories']}")


def page_alpha_access():
    st.markdown('<div class="am-section">Alpha → Access</div>', unsafe_allow_html=True)
    st.markdown('<div class="am-sub">Translate familiar practices into Access-aligned '
                'Guide moves.</div>', unsafe_allow_html=True)
    df = DATA["Alpha to Access"]
    for _, row in df.iterrows():
        card(row["Alpha_Practice"], row["Why_It_Works"], accent="#8b5cf6",
             icon_name="swap")
        with st.expander("Expand to Learn More"):
            st.markdown(f"**Future 2 constraint:** {row['Future_2_Constraint']}")
            st.markdown(f"**Access-aligned alternative:** {row['Access_Aligned_Alternative']}")


# ===========================================================================
# REFERENCE
# ===========================================================================
FOUNDATION_ICONS = ["compass", "users", "rocket", "bars", "loop", "target",
                    "network", "shield", "lifebuoy"]
FOUNDATION_COLORS = ["#3b82f6", "#6366f1", "#f59e0b", "#0ea5e9", "#ec4899",
                     "#8b5cf6", "#10b981", "#d946ef", "#14b8a6"]


def page_foundations():
    st.markdown('<div class="am-section">Foundations</div>', unsafe_allow_html=True)
    st.markdown('<div class="am-sub">The thinking behind the Access Model, drawn from '
                'the Access Model Brainlift. Expand any topic for deeper context.</div>',
                unsafe_allow_html=True)
    df = DATA["Brainlift References"]
    has_context = "Detailed_Context" in df.columns
    for i, (_, row) in enumerate(df.iterrows()):
        accent = FOUNDATION_COLORS[i % len(FOUNDATION_COLORS)]
        ic = FOUNDATION_ICONS[i % len(FOUNDATION_ICONS)]
        card(row["Reference_Topic"], row["Reference_Summary"], accent=accent, icon_name=ic)
        with st.expander("Expand to Learn More"):
            if has_context and str(row.get("Detailed_Context", "")).strip():
                st.markdown(f"**Deeper context**\n\n{row['Detailed_Context']}")
            if "Guide_Application" in df.columns and str(row.get("Guide_Application", "")).strip():
                st.markdown(f"**What this means for a Guide:** {row['Guide_Application']}")
            st.markdown(f"**Source:** {row['Source_Note']}")


def page_guide_role():
    st.markdown('<div class="am-section">Guide Role</div>', unsafe_allow_html=True)
    st.markdown('<div class="am-sub">The seven Access Model principles through what a '
                'Guide looks for and what strong implementation looks like.</div>',
                unsafe_allow_html=True)

    st.markdown(support_expectations_matrix(), unsafe_allow_html=True)
    st.caption("Rescuing and pressure are the common traps. The Guide's craft is "
               "staying in the high support, high expectations quadrant.")

    for _, row in PRINCIPLES.iterrows():
        principle = str(row["Principle"]).strip()
        card(principle, row["Plain_Language_Definition"], accent="#6366f1",
             icon_name=principle_icon(principle))
        with st.expander("Expand to Learn More"):
            st.markdown(f"**What the Guide looks for:** {row['Guide_Look_For']}")
            st.markdown(f"**Strong implementation:** {row['Strong_Implementation']}")
            st.markdown(f"**Common trap:** {row['Common_Misinterpretation']}")

    st.markdown('<div class="am-section" style="font-size:1.1rem;">Principles across '
                'the scenario bank</div>', unsafe_allow_html=True)
    st.caption("How often each principle anchors a scenario in this training set.")
    prin_counts = (SCEN["Access_Model_Principle"].astype(str)
                   .value_counts().rename_axis("Principle").to_frame("Scenarios"))
    st.bar_chart(prin_counts, horizontal=True, color="#6366f1")


# ===========================================================================
# Router
# ===========================================================================
PAGES = {
    "Session Progress": page_dashboard,
    "Missions": page_missions,
    "Guide Certification": page_certification,
    "Scenario Challenge": page_practice,
    "Launch Toolkit": page_launch_toolkit,
    "Check Charts": page_check_charts,
    "Repair Protocols": page_repair,
    "Alpha → Access": page_alpha_access,
    "Foundations": page_foundations,
    "Guide Role": page_guide_role,
    "Missions Completed": page_missions_completed,
    "Scenarios Completed": page_scenarios_completed,
}

render_sidebar()
PAGES.get(st.session_state["nav"], page_dashboard)()
