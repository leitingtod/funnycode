    public static String getLongestCommonPrefix(String[] strs) {
    if (strs.length == 0) {
      return "";
    }
    String ans = strs[0];
    for (int i = 1; i < strs.length; i++) {
      int j = 0;
      for (; j < ans.length() && j < strs[i].length(); j++) {
        if (ans.charAt(j) != strs[i].charAt(j)) {
          break;
        }
      }
      ans = ans.substring(0, j);
      if (ans.equals("")) {
        return ans;
      }
    }
    return ans;
  }
  
  public static String getMaxSubString(String s1, String s2) {
    String max, min;
    max = s1.length() > s2.length() ? s1 : s2;
    min = max.equals(s1) ? s2 : s1;
    for (int x = 0; x < min.length(); x++) {
      for (int y = 0, z = min.length() - x; z != min.length() + 1; y++, z++) {
        String temp = min.substring(y, z);
        if (max.contains(temp)) {
          return temp;
        }
      }
    }
    return "";
  }

  public static String capFirstLetter(String name) {
    char[] cs = name.toCharArray();
    cs[0] -= 32;
    return String.valueOf(cs);
  }
  
    public static void write(Path path, String s) {
    try {
      new File(path.getParent().toString()).mkdirs();
      Files.write(path, s.getBytes());
    } catch (IOException e) {
      e.printStackTrace();
    }
  }

  public static void write(Path path, String s, OpenOption... options) {
    try {
      new File(path.getParent().toString()).mkdirs();
      Files.write(path, s.getBytes(), options);
    } catch (IOException e) {
      e.printStackTrace();
    }
  }
  
  private static Comparator<Entry> comparatorByKeyAsc =
      (Entry o1, Entry o2) -> {
        if (o1.getKey() instanceof Comparable) {
          return ((Comparable) o1.getKey()).compareTo(o2.getKey());
        }
        throw new UnsupportedOperationException("键的类型尚未实现Comparable接口");
      };

  private static Comparator<Entry> comparatorByKeyDesc =
      (Entry o1, Entry o2) -> {
        if (o1.getKey() instanceof Comparable) {
          return ((Comparable) o2.getKey()).compareTo(o1.getKey());
        }
        throw new UnsupportedOperationException("键的类型尚未实现Comparable接口");
      };

  private static Comparator<Entry> comparatorByValueAsc =
      (Entry o1, Entry o2) -> {
        if (o1.getValue() instanceof Comparable) {
          return ((Comparable) o1.getValue()).compareTo(o2.getValue());
        }
        throw new UnsupportedOperationException("值的类型尚未实现Comparable接口");
      };

  private static Comparator<Entry> comparatorByValueDesc =
      (Entry o1, Entry o2) -> {
        if (o1.getValue() instanceof Comparable) {
          return ((Comparable) o2.getValue()).compareTo(o1.getValue());
        }
        throw new UnsupportedOperationException("值的类型尚未实现Comparable接口");
      };

  /** 按键升序排列 */
  public static <K, V> Map<K, V> sortByKeyAsc(Map<K, V> originMap) {
    if (originMap == null) {
      return null;
    }
    return sort(originMap, comparatorByKeyAsc);
  }

  /** 按键降序排列 */
  public static <K, V> Map<K, V> sortByKeyDesc(Map<K, V> originMap) {
    if (originMap == null) {
      return null;
    }
    return sort(originMap, comparatorByKeyDesc);
  }

  /** 按值升序排列 */
  public static <K, V> Map<K, V> sortByValueAsc(Map<K, V> originMap) {
    if (originMap == null) {
      return null;
    }
    return sort(originMap, comparatorByValueAsc);
  }

  /** 按值降序排列 */
  public static <K, V> Map<K, V> sortByValueDesc(Map<K, V> originMap) {
    if (originMap == null) {
      return null;
    }
    return sort(originMap, comparatorByValueDesc);
  }

  private static <K, V> Map<K, V> sort(Map<K, V> originMap, Comparator<Entry> comparator) {
    return originMap.entrySet().stream()
        .sorted(comparator)
        .collect(
            Collectors.toMap(Entry::getKey, Entry::getValue, (e1, e2) -> e2, LinkedHashMap::new));
  }
  
  public static List<Field> sortWithOrderList(List<Field> fieldList, List<String> orderList) {
    Map<String, Field> fieldMap = new LinkedHashMap<>();
    for (Field field : fieldList) {
      fieldMap.put(field.getName(), field);
    }
    List<Field> fields = new ArrayList<>();
    for (String s : orderList) {
      fields.add(fieldMap.get(s));
    }
    return fields;
  }

  public static List<Field> getObjectFields(Class<?> clazz) {
    List<Field> filedList = new ArrayList<>();

    while (clazz != null) {
      filedList.addAll(Arrays.asList(clazz.getDeclaredFields()));
      clazz = clazz.getSuperclass();
    }
    return filedList;
  }
  
    public static String pathShortener(String path) {
    return pathShortener(path, 3);
  }

  public static String pathShortener(String path, int threshold) {
    List<String> pathSplits = Splitter.on("\\").omitEmptyStrings().splitToList(path);

    if (pathSplits.size() > threshold) {
      return "..\\"
          + Joiner.on("\\")
          .join(pathSplits.subList(pathSplits.size() - threshold, pathSplits.size()));
    }

    return path;
  }
  
    public static boolean isChinese(char c) {
    Character.UnicodeScript sc = Character.UnicodeScript.of(c);
    if (sc == Character.UnicodeScript.HAN) {
      return true;
    }
    return false;
  }
  
    public static String getFirstEnglishWord(String s) {
    int start = 0, end = 0;
    char[] chars = s.toCharArray();
    for (int i = 0; i < chars.length; i++) {
      if (Character.isLetterOrDigit(chars[i])) {
        start = i;
        break;
      }
    }
    for (int i = start; i < chars.length; i++) {
      //      System.out.println(
      //          chars[i]
      //              + " isLetterOrDigit= "
      //              + Character.isLetterOrDigit(chars[i])
      //              + ", isChinese= "
      //              + Utils.isChinese(chars[i]));
      if (!Character.isLetterOrDigit(chars[i]) || Utils.isChinese(chars[i])) {
        end = i;
        break;
      }
    }
    if (end == 0) {
      end = s.length();
    }

    //    System.out.println("DefName = " + s.substring(start, end));
    return s.substring(start, end).trim();
  }
