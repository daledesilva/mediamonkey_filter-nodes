# Filter Nodes

A MediaMonkey script that simplifies the complexity of the library selection nodes.

---

Do you not need the complexity of the Library node and just want to see all your music by clicking on the first node (and navigate with the trackbrowser/search like in iTunes)? Do you want to set up your filters but then be able to access them by having a separate node for each?

The new filters in MediaMonkey 3 work well. They let you define subsets of music that you would like to group together when browsing (not unlike separate databases). The only problem is that you have to right click on the Library node each time you want to see a new filter.

**Filter Nodes creates main nodes with the same names as your filters.** Simply click on these nodes and MediaMonkey will activate that filter and show you all tracks associated with it.

- You don't need to expand the node to then click on the "Artist" node.
- Because each view uses a different filter, it remembers your view settings, so each of these nodes can have different columns set up!

**Setting up Subnodes:**

If one of your filters can be divided up into more categories, you can create subnodes off them by setting up another filter with a certain naming convention:

If your main node is named `My Music`, then you can create a subnode by naming another filter `My Music -- subnode name`.

*(The important thing is to divide it with **2 consecutive hyphens**)*

**Using Keywords for assistance:**

Filter Nodes also gives some assistance with certain keywords. At the moment, these include "Albums" and "Non-Albums".

Extending on the example above, you could name a filter `My Music -- Albums` (giving it the same criteria as My Music) and it would only show songs in which the Album Name is defined.

*(The reason this has been included is because currently, it is complicated in MM3 to create a filter which does this)*

The keyword "Non-Albums" doesn't restrict the tracks, it simply makes the "default" sort order Artist then Title.

(Restricting the tracks to only songs without albums can be done by including the criteria within your filter: `Album is Unknown`)

**Options** (accessed through tools/options/"Filter Nodes Settings"):

- Hide / unhide the Library.
- Auto expand playlist node on startup.
- Expand main filter nodes on startup.
- Alternative icons for now playing & library.
- Keyword assistance.

Hope you all like the script.

Dale.

Bugs:
- Currently, the track browser doesn't work properly with "script created nodes" and the album art views - this bug has already been fixed and should affect upcoming MM releases.


Community Forum: https://www.mediamonkey.com/forum/viewtopic.php?t=27303
