# Obsidian Vault Healthcheck v1.0
## Produces a Spreadsheet/Workbook of all **Properties**, their **Values**, and **Tags**
This system will run a set of python scripts and read[^1] through all Markdown files in a given vault and gather statistics on all **Properties**, their **Values**, and **Tags** (both frontmatter and inline), producing a spreadsheet for further user analysis. It will also document all duplicates filenames found in the vault, as well as any possibly corrupt YAML[^2]

This runs a set of python scripts that 

### Installation
1. Download/clone and unpack the repository to a folder
2. Run v_chk.py and fill in the following required values:
   1. The full path of the Obsidian Vault you wish to analyze
   2. The full pathname of your spreadsheet executable
   3. All other options can be left, as is.
3. Click on Save and Run and give it a few seconds to gather statistics
4. The spreadsheet will load automatically!
5. Consider buying me, a poor coder (steady there!), a coffee.

Note: I am in the process of learning how OOP works, as well as Python, so if you want to read/analyze the code, please be kind. My background in coding is limited and primarily functional. That having been said, this tool works well, at least for me.

[^1]: All scripts are READ-ONLY. This will not make any changes to your vault!
[^2]: Corrupt is defined as whatever PyYAML 6.0.2 is unable to safe_load.
