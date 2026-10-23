Real Numbers files used by the tests, taken from the test suite of
[numbers-parser](https://github.com/masaccio/numbers-parser) (MIT licence, see `LICENSE-numbers-parser.rst`).

`images.numbers` is `test-comments.numbers` with pictures, a group and shapes added by
[`tools/make_images_fixture.py`](../../tools/make_images_fixture.py), as none of the samples from Numbers has pictures.

The UI smoke test opens every `.numbers` file here and in folders under it, so new files (say, ones saved by a newer
Numbers) can simply be dropped in a folder here; [`test/ui/expectations.js`](../ui/expectations.js) says what
particular files must show.
