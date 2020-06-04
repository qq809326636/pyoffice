#!/usr/bin/env bash

# Get current folder path
currFolder=$(cd `dirname $0`; pwd)
pushd $currFolder

# Clean old rst
if [ -d $currFolder/source ]
then
  rm -f $currFolder/source/pyoffice*
  rm -f $currFolder/source/modules.rst
fi

# Create rst
sphinx-apidoc -o $currFolder/source $currFolder/../src/pyoffice
if [ $? -eq 0 ]
then
  make html
  if [ $? -eq 0 ]
  then
    rm -f $currFolder/source/pyoffice*
    rm -f $currFolder/source/modules.rst
    rm -rf $currFolder/build/doctrees
  else
    echo "Make html failed."
    for i in $(ls $currFolder/build)
    do
      rm -rf $currFolder/build/$i
    done
  fi
else
  echo "Create rst failed."
fi

popd
echo "End"
