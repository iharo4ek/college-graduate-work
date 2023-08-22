namespace Medplast{
    class User{
        private static User instance;
        private int id;
        private string login = "";
        private string password = "";
        private string jobTitle = "";
        private string sName = "";
        private string name = "";
        private string p = "";
        private User() { }
        public static User getInstance() {
            if (instance == null)
                instance = new User();
            return instance;
        }
        public string getLogin() {
            return this.login;
        }
        public void setLogin(string login) {
            this.login = login;
        }
        public string getPassword(){
            return this.password;
        }
        public void setPassword(string password) {
            this.password= password;
        }
        public string getJobTitle(){
            return this.jobTitle;
        }
        public void setJobTitle(string jobTitle) {
            this.jobTitle = jobTitle;
        }
        public int getId() {
            return this.id;
        }
        public void setId(int id) {
            this.id = id;
        }
        public string getName() {
            return this.name;
        }
        public void setName(string name) {
            this.name = name;
        }
        public string getSName() {
            return this.sName;
        }
        public void setSName(string name) {
            this.sName = name;
        }
        public string getP() {
            return this.p;
        }
        public void setP(string p) {
            this.p = p;
        }
    }
}
