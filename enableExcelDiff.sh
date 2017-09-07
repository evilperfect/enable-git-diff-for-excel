#!/bin/bash

gitignoreFile=".gitignore"
gitAttributesFile=".gitattributes"
gitConfigFile=".git/config"
excel2txtPath="excel2txtTool/excel2TxtTools.py"
excel2txtCmd="python $excel2txtPath"

#gitattributes
if [ ! -f "$gitAttributesFile" ]; then
 touch "$gitAttributesFile"
 echo "*.xlsx diff=excel2txt_diff_tool" >> $gitAttributesFile
 echo "*.xls diff=excel2txt_diff_tool" >> $gitAttributesFile
else
  grep -q "*.xlsx diff=excel2txt_diff_tool" $gitAttributesFile
  if [ $? -ne 0 ];then
    echo "*.xlsx diff=excel2txt_diff_tool" >> $gitAttributesFile
  fi
  grep -q "*.xls diff=excel2txt_diff_tool" $gitAttributesFile
  if [ $? -ne 0 ];then
    echo "*.xls diff=excel2txt_diff_tool" >> $gitAttributesFile
  fi
fi

#git config file
if [ -f "$gitConfigFile" ]; then
  grep -q "\[diff \"excel2txt_diff_tool\"\]" $gitConfigFile
  if [ $? -ne 0 ]
  then
    echo " " >> $gitConfigFile
    echo "[diff \"excel2txt_diff_tool\"]" >> $gitConfigFile
    echo "    binary = true" >> $gitConfigFile
    echo "    textconv = ${excel2txtCmd}" >> $gitConfigFile
  fi
fi

echo "excel2txt has been configured for git sucessfully. Making human-readable 'git diff' output for Excel files(*.xlsx or *.xls) possible."
