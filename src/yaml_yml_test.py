import yaml

def clean_links(yml):
    # This is very brute force, but I haven't got  a better way right now
    yml = f"{yml}"  # make sure it's a string this will destroy lists!!!
    yml = yml.replace("{'", "")
    yml = yml.replace("'}", "")
    yml = yml.replace("['", "[")      # remove inner single quotes                        [['test']]
    yml = yml.replace("']", "]")
    yml = yml.replace('["', '[')      # remove inner double quotes                         [["test"]]
    yml = yml.replace('"]', ']')
    yml = yml.replace('"[', '[')      # Cleanup this scenario:                            ["[test]"]"
    yml = yml.replace(']"', ']')      #   which will result in:                           [[test]]
    # At the time of this writing, there are no "lists" of links, but just in case
    # the linter created one, it would likely be like [[[one link]]] and I don't want that!
    yml = yml.replace("[[[", "[[")    # Just in case, clean up List of links              [[[test]]]
    yml = yml.replace("]]]", "]]")    #   which will result in:                           [[test]]
    # lastly, let's forces quotes around all [[]]
    yml = yml.replace("[[", '"[[')      # remove inner single quotes                        [['test']]
    yml = yml.replace("]]", ']]"')
    return yml

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
    print(f'{k}: <{v}> ({type(v)})')
    z = clean_links(v)
    print(f'{k}: <{z}> ({type(z)})\n')

def fix_yml_wikilinks(v):
    pass