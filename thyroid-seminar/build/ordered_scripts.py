import re
S={}
S['So, after treatment, how do these patients actually do?']='''SCRIPT (about 2 minutes; follows the slide left → middle → right → banner)
"Now we move to prognosis and outcomes. This slide has three parts: follow-up, prognostic factors, and survival.

[left column] On the left is follow-up in differentiated thyroid cancer, meaning papillary and follicular. After surgery we follow three things: thyroglobulin, anti-thyroglobulin antibodies, and a neck ultrasound. Thyroglobulin is made only by thyroid tissue, so after the gland is removed it should be very low. If it rises, some thyroid tissue, usually cancer, is still there. We check the antibodies with it because they can make the result falsely low.
At six to twelve months, the patient is placed in one of four response groups. Excellent: no evidence of disease, and the recurrence risk is only about one to four percent. Indeterminate: the results are unclear, and most of these patients stay stable. Biochemical incomplete: thyroglobulin is raised, but imaging is negative. Structural incomplete: disease is visible on imaging, and this is the highest-risk group.

[middle column] In the middle are the prognostic factors. Age is the most important, with fifty-five years as the AJCC cut-off. Distant metastases are the worst factor. For the tumour itself, gross extension outside the gland and aggressive histology worsen the outcome, and so does an incomplete resection.

[right column] On the right is five-year survival from SEER. It includes all thyroid cancers, but most of them are differentiated. Survival is almost one hundred percent for localized and regional disease, and about fifty percent with distant spread.
Below the chart: medullary cancer is followed with calcitonin and CEA instead of thyroglobulin, because it comes from C cells. Its ten-year survival falls from ninety-six percent when localized to forty percent with distant spread. Anaplastic cancer has a median survival of only about five months.

[banner] So the key message is: thyroid cancer rarely causes death, but it often recurs, mostly in the neck nodes within the first ten years. That is why follow-up is long."
'''
S['Outcome also includes what the surgery itself can cause.']='''SCRIPT (about 1.5 minutes; follows the slide left → middle → right → banner)
"The next slide looks at surgical outcomes: the complications of thyroid surgery, and how surgery affects cancer control.

[left column] On the left is nerve injury. The recurrent laryngeal nerve moves the vocal cord, so injury causes hoarseness. It is temporary in three to eight percent of patients and permanent in less than one to two percent. If both nerves are injured, both vocal cords can close, and the airway is at risk. The external branch of the superior laryngeal nerve controls voice pitch, so its injury causes loss of the high-pitched voice and voice fatigue.

[middle column] In the middle are metabolic and airway complications. The parathyroid glands lie behind the thyroid and can be injured or lose their blood supply during surgery. This causes low calcium, the most common complication, in ten to thirty percent of patients. It is usually temporary, and permanent in less than one to two percent. A neck hematoma is rare, about one percent, but it can compress the airway, so it needs immediate evacuation.

[right column] On the right is oncologic control. An R0 resection means the tumour is completely removed, with nothing left behind. It is the strongest factor the surgeon controls to prevent recurrence. After surgery, the ATA risk group guides two decisions: whether the patient needs radioactive iodine, and how much we suppress TSH.

[banner] Finally, complication rates are clearly lower when the surgeon performs more than about twenty-five thyroid operations a year."
'''
S["Let's apply this with three quick cases."]='''SCRIPT (about 2 minutes, interactive; 9 clicks in the order Scenario → Think → Decide)
"Finally, let's apply this to three cases. For each one, I'll show the scenario first, give you a moment to think, and then show the answer.

[click 1: Case 1 Scenario] Case one is a thirty-four-year-old woman with a two-centimetre nodule. The ultrasound is TI-RADS five, which is highly suspicious, and the FNA is Bethesda six, which means malignant. What is the diagnosis, and what is her risk?
[pause, then click 2: Think] This is papillary cancer. Because she is under fifty-five, she is stage one, and her risk is low.
[click 3: Decide] She had a thyroidectomy without routine radioactive iodine. One year later her thyroglobulin is undetectable, which is an excellent response.

[click 4: Case 2 Scenario] Case two is a forty-two-year-old man with a thyroid nodule and a neck node. He has diarrhoea, and his father had thyroid cancer and episodes of severe high blood pressure. What should we suspect?
[pause, then click 5: Think] Medullary cancer, most likely as part of MEN2A.
[click 6: Decide] Before surgery we must exclude a pheochromocytoma, because operating without knowing about it can cause a dangerous rise in blood pressure. Then we do a total thyroidectomy with neck dissection, and we follow him with calcitonin and CEA.

[click 7: Case 3 Scenario] Case three is a seventy-four-year-old woman with a hard neck mass that has grown over three weeks. She is hoarse and short of breath. What are we worried about?
[pause, then click 8: Think] Anaplastic cancer, but we must also consider lymphoma, because lymphoma is treatable. So we take a core biopsy and test for BRAF.
[click 9: Decide] Management starts with securing the airway, then an urgent multidisciplinary discussion, BRAF-targeted therapy if the test is positive, and early palliative care.

These three cases show how different thyroid cancers can be. Thank you."
'''
p='build.js'; s=open(p).read()
for start,new in S.items():
    i=s.index(start); a=s.rindex('SCRIPT (',0,i); b=s.index('BACKUP Q&A',i)
    s=s[:a]+new+'\n'+s[b:]
open(p,'w').write(s)
tot=0
for v in S.values():
    w=len(re.findall(r"[A-Za-z0-9'-]+",re.sub(r'\[[^\]]*\]','',v.split('\n',1)[1]))); tot+=w; print(w)
print('total',tot)
