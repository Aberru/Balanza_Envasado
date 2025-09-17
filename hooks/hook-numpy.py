# Hook for numpy to fix PyInstaller import issues
from PyInstaller.utils.hooks import collect_data_files, collect_submodules, collect_dynamic_libs

# Collect all numpy submodules
hiddenimports = collect_submodules('numpy')

# Collect numpy data files
datas = collect_data_files('numpy')

# Collect numpy dynamic libraries
binaries = collect_dynamic_libs('numpy')

# Add specific numpy modules that are commonly missed
hiddenimports += [
    'numpy.core.multiarray',
    'numpy.core.umath',
    'numpy.core._dtype_ctypes',
    'numpy.core._internal',
    'numpy.core._exceptions',
    'numpy.core._type_aliases',
    'numpy.core._string_helpers',
    'numpy.core._asarray',
    'numpy.core._multiarray_umath',
    'numpy.core._multiarray_tests',
    'numpy.core._operand_flag_tests',
    'numpy.core._struct_ufunc_tests',
    'numpy.core._umath_tests',
    'numpy.core._rational_tests',
    'numpy.core._umath_linalg',
    'numpy.lib.utils',
    'numpy.lib.arraysetops',
    'numpy.lib.npyio',
    'numpy.lib.format',
    'numpy.lib.mixins',
    'numpy.lib.scimath',
    'numpy.lib.stride_tricks',
    'numpy.lib.twodim_base',
    'numpy.lib.type_check',
    'numpy.lib.ufunclike',
    'numpy.lib.arraypad',
    'numpy.lib.arrayterator',
    'numpy.lib.function_base',
    'numpy.lib.histograms',
    'numpy.lib.index_tricks',
    'numpy.lib.nanfunctions',
    'numpy.lib.polynomial',
    'numpy.lib.shape_base',
]
