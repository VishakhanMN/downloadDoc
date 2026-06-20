/**
 * Custom esbuild plugin to mark bootstrap.min.css as external.
 * This prevents esbuild from trying to bundle the file and leaves
 * the unresolved path as-is in the output.
 */
export default [
  {
    name: 'bootstrap-external',
    setup(build) {
      build.onResolve(
        { filter: /bootstrap\/dist\/css\/bootstrap\.min\.css$/ },
        () => ({
          path: 'node_modules/bootstrap/dist/css/bootstrap.min.css',
          external: true,
        })
      );
    },
  },
];
