{
  'targets': [
    {
      'target_name': 'lib_node_activex',
      'type': 'static_library',
      'sources': [
        'src/utils.cpp',
        'src/disp.cpp'
      ],
      'defines': [
        'BUILDING_NODE_EXTENSION',
      ],
      'direct_dependent_settings': {
        'include_dirs': ['include']
      },
      'dependencies': [
      ]
    }
  ]
}
