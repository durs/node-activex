{
  'targets': [
    {
      'target_name': 'node_activex',
      'sources': [
        'src/main.cpp',
        'src/utils.cpp',
        'src/disp.cpp'
      ],
      'libraries': [ 
      ],
      'dependencies': [      
      ]
    },
    {
      'target_name': 'lib_node_activex',
      'type': 'static_library',
      'sources': [
        'src/utils.cpp',
        'src/disp.cpp'
      ],
      'direct_dependent_settings': {
        'include_dirs': ['.']
      },
      'dependencies': [
      ]
    }
  ]
}