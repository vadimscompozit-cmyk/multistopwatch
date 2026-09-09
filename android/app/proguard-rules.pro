# Мост в JavaScript вызывается по именам методов — их нельзя переименовывать.
-keepclassmembers class com.sasin.multistopwatch.MainActivity$Bridge {
    public *;
}
