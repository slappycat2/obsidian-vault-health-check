import yaml

y = """  
test1: [[🗺️ Normal Wikilink]]
test2: "[This is a markdown link](This%20is%20a%20markdown%20link.md)"
test3: ["This", "list", "is", "ok"]  
test5: "[[Normal Wikilink string]]"  
test6: [['How should this work?']]
test7: 
test8:  
 - [[🗺️ This link has an emoji]] 
 - [[⚒️ These have no quotes]] 
 - [[🗺️ but all are a-ok]]
"""

yml = yaml.load(y, Loader=yaml.SafeLoader)

print(f"\nBefore:{y}\nAfter PyYaml Loading:")
for k, v in yml.items():
    print(f'{k}: <{v}>" ({type(v)})')

def fix_yml_wikilinks(v):
    pass