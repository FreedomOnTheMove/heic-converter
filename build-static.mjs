import * as esbuild from 'esbuild'
import { copyFileSync, mkdirSync, existsSync } from 'fs'
import { join } from 'path'

// Create build directory
if (!existsSync('static-build')) {
    mkdirSync('static-build')
}

// Bundle the app with all dependencies
await esbuild.build({
    entryPoints: ['example/app.js'],
    bundle: true,
    format: 'esm',
    outfile: 'static-build/app.bundle.js',
    minify: true,
    target: 'es2020'
})

// Copy HTML and other assets
copyFileSync('example/index.html', 'static-build/index.html')
copyFileSync('example/jszip.min.js', 'static-build/jszip.min.js')

console.log('✅ Static build complete in ./static-build/')