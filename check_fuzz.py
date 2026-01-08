from thefuzz import fuzz

names = [
    ('Xomi', 'Xorii'),
    ('Ион', 'Йоныч'),
    ('Бомбилаат', 'Бомбилаа'),
    ('Самарешу', 'Самарешу'),
    ('Самарешу', 'Самрешу'),
]

print(f"{'Source':<15} | {'Target':<15} | {'Ratio':<10} | {'Partial Ratio':<15}")
print("-" * 60)
for s, t in names:
    r = fuzz.ratio(s, t)
    pr = fuzz.partial_ratio(s, t)
    print(f"{s:<15} | {t:<15} | {r:<10} | {pr:<15}")
