plugins {
    java
}
sourceSets {
    main {
        resources {
            srcDirs("src/assets", "src/assets-gen")
        }
        java {
            srcDir("src/main")
        }
    }
}
allprojects {
    repositories {
        mavenCentral()
    }
    dependencies {
        implementation("io.github.osobolev:jacob:1.20")
    }
}
tasks.withType<JavaExec>() {
    systemProperty("java.library.path", file("\\src\\assets"))
}