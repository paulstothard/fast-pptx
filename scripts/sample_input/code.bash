if [[ ! -e $target ]]; then
    mkdir -p $target
elif [[ ! -d $target ]]; then
    echo "$target already exists but is not a directory" 1>&2
fi
