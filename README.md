# test-xl-git

> testing using git for excel based VBA development projects.

## Features

- a [pre-commit hook (pre-commit.py)](res/pre-commit.py) using python to automatically extract all objects from the VBA project before each commit into the [src](src/) folder. 
    - See Also: the accompanying [pre-commit](res/pre-commit) script that calls [pre-commit.py](res/pre-commit.py).

- a [github release action](.github/workflows/release-xl.yml) to automatically build the VBA project on each github release.


