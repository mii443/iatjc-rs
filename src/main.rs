use iatjc_rs::tsf::TSF;
use iatjc_rs::com::Com;

fn main() {
    tracing_subscriber::fmt::init();

    let _com = Com::new().unwrap();

    let mut tsf_main = TSF::new();
    tsf_main.initialize().unwrap();

    println!("TSF initialized successfully");
}