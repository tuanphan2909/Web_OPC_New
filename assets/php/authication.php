<?php

use ~/model/Account;
class SampleController extends Controller
{

    public function index()
    {
        // Thực hiện xử lý và lấy thông tin tài khoản từ mô hình
        
        $username = Account::Find(1);
          // Giả sử tên đăng nhập của người dùng là "opcpharma"

        // Kiểm tra tên đăng nhập và gán vai trò tương ứng
        if ($username === 'admin') {
            $role = 'admin';
        } else {
            $role = 'user';
        }

        // Truyền giá trị vai trò vào biến ViewBag để sử dụng trong View
        $this->viewBag['role'] = $role;

        // Render view
        $this->render('index');
    }
}
?>
