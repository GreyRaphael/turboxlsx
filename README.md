# fastxlsx

high performance python library to write xlsx file by rust pyo3

## Development

```bash
# activate python environment
source ~/envs/jupy12/bin/activate
# install maturin
pip install --upgrade maturin

maturin init
# choose pyo3

# change Cargo.toml features to 
# features = ["abi3-py38"]

maturin develop

# begin release *whl
maturin build --release

# begin publish to pypi
maturin publish
```