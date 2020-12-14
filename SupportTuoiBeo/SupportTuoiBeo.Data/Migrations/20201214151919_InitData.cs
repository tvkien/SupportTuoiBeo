using Microsoft.EntityFrameworkCore.Migrations;

namespace SupportTuoiBeo.Data.Migrations
{
    public partial class InitData : Migration
    {
        protected override void Up(MigrationBuilder migrationBuilder)
        {
            migrationBuilder.CreateTable(
                name: "UserDetails",
                columns: table => new
                {
                    Id = table.Column<int>(type: "int", nullable: false)
                        .Annotation("SqlServer:Identity", "1, 1"),
                    MaKH = table.Column<string>(type: "nvarchar(450)", nullable: true),
                    Tinh = table.Column<string>(type: "nvarchar(max)", nullable: true),
                    TienThanhToan = table.Column<long>(type: "bigint", nullable: false),
                    Ngay = table.Column<int>(type: "int", nullable: false),
                    Thang = table.Column<int>(type: "int", nullable: false)
                },
                constraints: table =>
                {
                    table.PrimaryKey("PK_UserDetails", x => x.Id);
                });

            migrationBuilder.CreateIndex(
                name: "IX_UserDetails_MaKH",
                table: "UserDetails",
                column: "MaKH",
                unique: true,
                filter: "[MaKH] IS NOT NULL");
        }

        protected override void Down(MigrationBuilder migrationBuilder)
        {
            migrationBuilder.DropTable(
                name: "UserDetails");
        }
    }
}
