<?php
namespace app\controllers;

use Yii;
use yii\filters\AccessControl;
use yii\web\Controller;
use yii\web\Response;
use yii\filters\VerbFilter;
use app\models\LoginForm;
use app\models\ContactForm;

class ExcelController extends Controller
{
    /**
     * @name actionTest
     * @author fyh
     * @param $a
     * @return mixed
     */
    public function actionTest($a){
        $data = \moonland\phpexcel\Excel::import('/vagrant/basic/HPV.xlsx',[
            'setFirstRecordAsKeys' => false, // if you want to set the keys of record column with first record, if it not set, the header with use the alphabet column on excel.
            'setIndexSheetByName' => false, // set this if your excel data with multiple worksheet, the index of array will be set with the sheet name. If this not set, the index will use numeric.
//            'getOnlySheet' => 'sheet1', // you can set this property if you want to get the specified sheet from the excel data with multiple worksheet.
        ]);
        echo "<pre>";
        var_dump($data);
        echo "</pre>";
        return $a;


  
  

  


<?php
/**
 * Created by PhpStorm.
 * User: Administrator
 * Date: 2018/7/26
 * Time: 10:57
 */
namespace app\modules\orders\services;

use common\models\Consignee;
use common\models\Goods;
use common\models\Order;
use common\models\OrderConsignee;
use common\services\BaseService;
use yii\db\Exception;

class ImportService extends BaseService{
    public function importDispatch($data){ //处理单行数据
        $transaction = yii::$app->db->beginTransaction();
        try {
            $order_sn = $this->importOrder($data);//订单表处理
            $consignee_id=$this->importConsignee($data);//收货人信息表处理
            $order_consignee_id = $this->importOrderConsignee($order_sn,$consignee_id);//订单与收货人关联信息
            $goods_id = $this->importGoods($data);
            if ($order_sn && $consignee_id && $order_consignee_id && $goods_id ) {
                $transaction->commit();
                return true;
            } else {
                throw new Exception($data[self::getImportKey('order_sn')]);
            }
        } catch (Exception $e) {
            $transaction->rollBack();
            return $e->getMessage();
        }
    }

    public function importOrder($data){
        $order = new Order();
        $order->order_sn = (string)$data[self::getImportKey('order_sn')];
        $order->shop_id = 1;//暂时写死
        $order->order_price = $data[self::getImportKey('pay_price')];
        $order->order_mark = (int)$data[self::getImportKey('order_mark')];
        $order->last_sync_time = date('Y-m-d H:i:s', time());
        $order->order_time = date('Y-m-d H:i:s', $data[self::getImportKey('order_time')]);
        $order->order_status = 0;//订单状态 0 待审核
        $order->pay_price = $data[self::getImportKey('pay_price')];
        $order->order_fee = $data[self::getImportKey('order_fee')];
        $order->examined_time = '';
        $order->create_time = time();
        $order->update_time = '';
        $order->buyer_message = $data[self::getImportKey('buyer_message')];
        $order->cm_message = $data[self::getImportKey('cm_message')];
        $order->member_name = $data[self::getImportKey('member_name')];
        return $order->save()?$order->attributes['order_sn']:false;
    }

    public function importConsignee($data){
        $consignee = new Consignee();
        $consignee->name=$data[self::getImportKey('name')];
        $consignee->province='';
        $consignee->city='';
        $consignee->district=$data[self::getImportKey('district')];
        $consignee->address=$data[self::getImportKey('address')];
        $consignee->mobile=$data[self::getImportKey('mobile')];
        $consignee->create_time=time();
        return $consignee->save()?$consignee->attributes['id']:false;
    }

    public function importOrderConsignee($order_sn,$consignee_id){
        $order_consignee = new OrderConsignee();
        $order_consignee->order_sn=$order_sn;
        $order_consignee->consignee_id=$consignee_id;
        $order_consignee->create_time=time();
        return $order_consignee->save()?$order_consignee->attributes['id']:false;
    }

    /**
     * @name importGoods
     * @author fyh
     * @param $data
     * @return bool
     */
    public function importGoods($data){
        $goods = new Goods();
        $goods->order_sn=$data[self::getImportKey('order_sn')];
        $goods->goods_sn=$data[self::getImportKey('goods_sn')];
        $goods->goods_name=$data[self::getImportKey('goods_name')];
        $goods->goods_num=$data[self::getImportKey('goods_num')];
        $goods->goods_spec=$data[self::getImportKey('goods_spec')];
        $goods->goods_prices=$data[self::getImportKey('goods_prices')];
        $goods->create_time=time();
        return $goods->save()?$goods->attributes['id']:false;
    }

    /**
     * @name getImportKey
     * @author fyh
     * @param $key
     * @return mixed
     */
    public static function getImportKey($key){
        $titleField = [
            'order_sn' => '订单号',
            'order_mark' => '状态标记',
            'order_time' => '下单时间',
            'pay_time' => '付款时间',
            'pay_price' => '付款金额',
            'pay_way' => '支付方式',//支付方式
            'order_price' => '订单金额',//订单金额
            'order_fee' => '运费',
            'member_name' => '购买会员ID',
            'name' => '收货人姓名',
            'district' => '地区',
            'address' => '详细地址',
            'mobile' => '电话',
            'goods_sn' => '商品货号',
            'goods_name' => '名称',
            'goods_spec' => '规格',
            'goods_num' => '数量',
            'goods_prices' => '金额',
            'buyer_message' => '买家留言',
            'cm_message' => '客服备注',
        ];
        return $titleField[$key];
    }
}
  
  
  
  
  
  
  
  
  
  
  
  
  
  
    }


}