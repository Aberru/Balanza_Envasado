# Hook for pandas to fix PyInstaller import issues
from PyInstaller.utils.hooks import collect_data_files, collect_submodules, collect_dynamic_libs

# Collect all pandas submodules
hiddenimports = collect_submodules('pandas')

# Collect pandas data files
datas = collect_data_files('pandas')

# Collect pandas dynamic libraries
binaries = collect_dynamic_libs('pandas')

# Add specific pandas modules that are commonly missed
hiddenimports += [
    'pandas._libs.tslibs.timedeltas',
    'pandas._libs.tslibs.np_datetime',
    'pandas._libs.tslibs.nattype',
    'pandas._libs.reduction',
    'pandas._libs.hashtable',
    'pandas._libs.lib',
    'pandas._libs.parsers',
    'pandas._libs.writers',
    'pandas._libs.algos',
    'pandas._libs.groupby',
    'pandas._libs.hashing',
    'pandas._libs.index',
    'pandas._libs.indexing',
    'pandas._libs.internals',
    'pandas._libs.join',
    'pandas._libs.missing',
    'pandas._libs.ops',
    'pandas._libs.properties',
    'pandas._libs.reshape',
    'pandas._libs.sparse',
    'pandas._libs.testing',
    'pandas._libs.tslib',
    'pandas._libs.tslibs',
    'pandas._libs.tslibs.base',
    'pandas._libs.tslibs.ccalendar',
    'pandas._libs.tslibs.dtypes',
    'pandas._libs.tslibs.fields',
    'pandas._libs.tslibs.nattype',
    'pandas._libs.tslibs.np_datetime',
    'pandas._libs.tslibs.offsets',
    'pandas._libs.tslibs.parsing',
    'pandas._libs.tslibs.period',
    'pandas._libs.tslibs.resolution',
    'pandas._libs.tslibs.strptime',
    'pandas._libs.tslibs.timedeltas',
    'pandas._libs.tslibs.timestamps',
    'pandas._libs.tslibs.timezones',
    'pandas._libs.tslibs.tzconversion',
    'pandas._libs.tslibs.vectorized',
    'pandas._libs.window',
]
