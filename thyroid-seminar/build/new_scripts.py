import re
S={}
S['Prognosis in one slide.']='''SCRIPT (about 80 seconds — speak it, don't read it)
"Okay, so the patient has had their treatment. Now the real question is: how are they going to do, and how do we keep an eye on them?

For differentiated cancer, papillary and follicular, follow-up is actually quite logical. Once the thyroid is out, the only thing in the body that makes thyroglobulin is thyroid tissue. So if Tg starts climbing, something is still there. We always send it with the antibodies, because the antibodies can hide it and give us a falsely reassuring number. Add a neck ultrasound, and at around six to twelve months we put the patient into one of four groups. [point to the left column] Excellent response is what we want: recurrence only about one to four percent. Structural incomplete, where we can actually see disease, is the worst. The two in the middle basically mean: keep watching.

Medullary is different. It doesn't come from follicular cells, so Tg is useless; we follow calcitonin and CEA instead.

What predicts outcome? If I had to pick one thing, it's age, with fifty-five as the AJCC cut-off. Then distant spread, gross invasion, aggressive histology, and whether we got it all out.

[point to the chart] And the good news: localized and regional disease sit at almost a hundred percent. It only really drops with distant spread. Anaplastic is the exception, about five months.

So the take-home: patients rarely die of thyroid cancer, but it can come back, usually in the neck, and that's why we follow them for years."
'''
S["From a surgeon's point of view"]='''SCRIPT (about 70 seconds)
"Since we're a surgery department, I want to show the other side of outcome: what the operation itself can cost the patient.

There are two nerves we really worry about. The recurrent laryngeal nerve: if it's injured, the patient wakes up hoarse. Most of the time it's temporary, three to eight percent, and it recovers; permanent damage is under one to two percent. But if both sides are injured, it becomes an airway problem. The external branch of the superior laryngeal is less dramatic, but a teacher or a singer will tell you they've lost their high notes.

Then calcium. The parathyroids are tiny and sit right behind the thyroid, so after a total thyroidectomy it's very common for calcium to drop for a few days, in up to a third of patients. Permanent problems are rare.

And the one complication everyone has to remember: a neck hematoma. It's rare, about one percent, but it can close the airway, so you open the wound right there at the bedside. You don't wait for theatre.

On the cancer side, the most important thing the surgeon controls is getting it all out: an R0 resection. After that, the ATA risk group decides on radioiodine and how much we suppress TSH.

[point to the banner] And one thing I found really interesting: surgeons who do more than about twenty-five thyroidectomies a year have clearly fewer complications. Volume matters."
'''
S['Three short illustrative cases.']='''SCRIPT (about 80 seconds)
"Let me finish with three quick patients, just to bring it all together. They're illustrative cases.

First, a thirty-four-year-old woman with a two-centimetre nodule. The ultrasound looks worrying, TI-RADS five, and the FNA comes back Bethesda six: papillary cancer. She's under fifty-five, so stage one and low risk. She has her operation with no routine radioiodine, and a year later her thyroglobulin is undetectable. Excellent response. That's the typical, happy thyroid cancer story.

Second, a forty-two-year-old man: a nodule, a node in the neck, diarrhea, and his father had thyroid cancer and episodes of very high blood pressure. That should ring a bell: medullary cancer, probably MEN2A. And here's the trap. Before we touch the thyroid, we have to rule out a pheochromocytoma, or he can crash on the table. Then a total thyroidectomy with neck dissection, and we follow him with calcitonin and CEA.

Third, a seventy-four-year-old lady with a hard neck mass that grew in three weeks. She's hoarse and short of breath. We think anaplastic, but we have to rule out lymphoma first, because lymphoma is very treatable. So we take a core biopsy and check BRAF. Airway first, an urgent team discussion, and honest, early palliative care.

Three nodules, three completely different outcomes. Thank you, I'm happy to take any questions."
'''
p='build.js'; s=open(p).read()
for start,new in S.items():
    i=s.index(start); a=s.rindex('SCRIPT (',0,i); b=s.index('BACKUP Q&A',i)
    s=s[:a]+new+'\n'+s[b:]
open(p,'w').write(s)
for v in S.values():
    print(len(re.findall(r"[A-Za-z0-9'-]+",v.split('\n',1)[1])))
