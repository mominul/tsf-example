extern crate embed_resource;

fn main() {
    embed_resource::compile("src/TextService.rc", embed_resource::NONE);
}
